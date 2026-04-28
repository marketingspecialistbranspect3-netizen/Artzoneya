"""Scan a folder of paintings, group duplicates/mockups, emit an Excel sheet.

Usage:
    python scan_paintings.py INPUT_DIR [-o paintings.xlsx] [--link-map links.csv]
                              [--phash-threshold 6] [--orb-min-matches 25]
                              [--no-feature-match]

Folder layout expected:
    INPUT_DIR/
        Abstract/
            painting1.jpg
            painting1_mockup.jpg     # framed-on-wall mockup
            painting2.png
            ...
        Portraits/
            ...

Output xlsx columns:
    Category | Painting name | Short description | Long description |
    Technical description | Picture link | Primary file | Mockup files |
    Mockup links | Notes

The "name" and three "description" columns are intentionally left blank for
manual fill-in. Picture link / Mockup links are filled if you pass a
--link-map CSV exported from Drive (see drive_link_export.gs).
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from PIL import Image
import imagehash
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import cv2  # type: ignore
    import numpy as np  # type: ignore
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".gif"}

# Filename hints that an image is a mockup / context shot rather than the
# bare painting. Used as a tie-breaker when picking the primary image of a
# duplicate cluster.
MOCKUP_KEYWORDS = (
    "mockup", "mock-up", "mock_up", "frame", "framed", "wall",
    "scene", "room", "interior", "context", "lifestyle",
)
THUMBNAIL_KEYWORDS = ("thumb", "thumbnail", "_sm", "small", "preview", "lowres", "low_res")


@dataclass
class ImageRecord:
    path: Path
    category: str
    phash: imagehash.ImageHash
    width: int
    height: int
    filesize: int
    cluster_id: int = -1
    is_primary: bool = False
    descriptors: object = field(default=None, repr=False)  # ORB descriptors

    @property
    def filename(self) -> str:
        return self.path.name

    @property
    def looks_like_mockup(self) -> bool:
        name = self.filename.lower()
        return any(k in name for k in MOCKUP_KEYWORDS)

    @property
    def looks_like_thumbnail(self) -> bool:
        name = self.filename.lower()
        return any(k in name for k in THUMBNAIL_KEYWORDS)

    @property
    def pixel_area(self) -> int:
        return self.width * self.height

    @property
    def aspect_ratio(self) -> float:
        return self.width / self.height if self.height else 0.0


def iter_image_files(root: Path) -> Iterable[tuple[Path, str]]:
    """Yield (path, category) for every image under root.

    Top-level subdirectories of root are treated as categories. Images
    directly inside root get category "(uncategorized)". Nested folders
    inside a category fold up to that category.
    """
    for entry in sorted(root.iterdir()):
        if entry.is_file() and entry.suffix.lower() in IMAGE_EXTS:
            yield entry, "(uncategorized)"
        elif entry.is_dir():
            category = entry.name
            for path in sorted(entry.rglob("*")):
                if path.is_file() and path.suffix.lower() in IMAGE_EXTS:
                    yield path, category


def load_record(path: Path, category: str, compute_orb: bool) -> ImageRecord | None:
    try:
        with Image.open(path) as im:
            im.load()
            phash = imagehash.phash(im, hash_size=16)
            width, height = im.size
    except Exception as exc:
        print(f"  ! skip {path}: {exc}", file=sys.stderr)
        return None

    descriptors = None
    if compute_orb:
        descriptors = compute_orb_descriptors(path)

    return ImageRecord(
        path=path,
        category=category,
        phash=phash,
        width=width,
        height=height,
        filesize=path.stat().st_size,
        descriptors=descriptors,
    )


def compute_orb_descriptors(path: Path, max_features: int = 1500):
    """Read an image, downscale, and compute ORB descriptors. Returns None on failure."""
    if not HAS_CV2:
        return None
    img = cv2.imread(str(path), cv2.IMREAD_GRAYSCALE)
    if img is None:
        return None
    h, w = img.shape
    scale = 800 / max(h, w)
    if scale < 1.0:
        img = cv2.resize(img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
    orb = cv2.ORB_create(nfeatures=max_features)
    _, desc = orb.detectAndCompute(img, None)
    return desc


def cluster_by_phash(records: list[ImageRecord], threshold: int) -> None:
    """Greedy single-link clustering on Hamming distance of pHash. Mutates records in place."""
    next_id = 0
    for rec in records:
        if rec.cluster_id != -1:
            continue
        rec.cluster_id = next_id
        for other in records:
            if other.cluster_id != -1:
                continue
            if (rec.phash - other.phash) <= threshold:
                other.cluster_id = next_id
        next_id += 1


def merge_clusters_by_orb(records: list[ImageRecord], min_matches: int) -> None:
    """Merge clusters where one painting appears embedded inside another (mockup).

    For each pair of distinct clusters, take one representative from each and
    run an ORB descriptor match. If enough good matches survive Lowe's ratio
    test, merge the clusters.
    """
    if not HAS_CV2:
        return

    bf = cv2.BFMatcher(cv2.NORM_HAMMING)

    # Pick the highest-resolution member of each cluster as its ORB
    # representative -- thumbnails have weak descriptors and miss matches.
    by_cluster: dict[int, ImageRecord] = {}
    for rec in records:
        if rec.descriptors is None:
            continue
        cur = by_cluster.get(rec.cluster_id)
        if cur is None or rec.pixel_area > cur.pixel_area:
            by_cluster[rec.cluster_id] = rec

    cluster_ids = list(by_cluster.keys())

    # Union-find for cluster merging
    parent = {cid: cid for cid in cluster_ids}

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    def good_pairs(des_a, des_b, ratio: float = 0.7) -> set[tuple[int, int]]:
        try:
            raw = bf.knnMatch(des_a, des_b, k=2)
        except cv2.error:
            return set()
        out: set[tuple[int, int]] = set()
        for pair in raw:
            if len(pair) < 2:
                continue
            m, n = pair
            if m.distance < ratio * n.distance:
                out.add((m.queryIdx, m.trainIdx))
        return out

    for i, ca in enumerate(cluster_ids):
        ra = by_cluster[ca]
        if ra.descriptors is None or len(ra.descriptors) < min_matches:
            continue
        for cb in cluster_ids[i + 1:]:
            if find(ca) == find(cb):
                continue
            rb = by_cluster[cb]
            if rb.descriptors is None or len(rb.descriptors) < min_matches:
                continue
            ab = good_pairs(ra.descriptors, rb.descriptors)
            if len(ab) < min_matches:
                continue
            ba = good_pairs(rb.descriptors, ra.descriptors)
            # Symmetric matches: a->b and b->a both pick the same pair.
            symmetric = sum(1 for q, t in ab if (t, q) in ba)
            if symmetric >= min_matches:
                union(ca, cb)

    # Apply union-find result
    for rec in records:
        rec.cluster_id = find(rec.cluster_id)


def pick_primaries(records: list[ImageRecord]) -> None:
    """For each cluster, mark exactly one record as the primary image."""
    by_cluster: dict[int, list[ImageRecord]] = defaultdict(list)
    for rec in records:
        by_cluster[rec.cluster_id].append(rec)

    for group in by_cluster.values():
        def score(r: ImageRecord) -> tuple:
            ar = r.aspect_ratio
            ar_penalty = 0.0 if 0.5 <= ar <= 2.0 else abs(ar - 1.0)
            return (
                1 if r.looks_like_mockup else 0,    # mockups last
                1 if r.looks_like_thumbnail else 0, # thumbnails last
                ar_penalty,                          # weird aspect ratios last
                -r.pixel_area,                       # higher resolution first
                r.filename.lower(),
            )
        group.sort(key=score)
        group[0].is_primary = True


def slugify(name: str) -> str:
    name = re.sub(r"\.[^.]+$", "", name)
    name = re.sub(r"[_\-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def load_link_map(path: Path) -> dict[str, str]:
    """Load a CSV of filename,url. Filenames are matched case-insensitively."""
    mapping: dict[str, str] = {}
    with path.open(newline="", encoding="utf-8") as fh:
        reader = csv.reader(fh)
        for row in reader:
            if len(row) < 2:
                continue
            fname, url = row[0].strip(), row[1].strip()
            if not fname or fname.lower() == "filename":
                continue
            mapping[fname.lower()] = url
    return mapping


SHEET_HEADERS = [
    "link",
    "name",
    "short description",
    "long description",
    "technical description",
]
SHEET_COL_WIDTHS = [40, 28, 40, 60, 40]


# Excel sheet name limits: <=31 chars, no [ ] : * ? / \
_INVALID_SHEET_CHARS = re.compile(r"[\[\]:\*\?/\\]")


def safe_sheet_name(name: str, taken: set[str]) -> str:
    cleaned = _INVALID_SHEET_CHARS.sub(" ", name).strip() or "Category"
    cleaned = cleaned[:31]
    base = cleaned
    i = 2
    while cleaned.lower() in {t.lower() for t in taken}:
        suffix = f" {i}"
        cleaned = base[: 31 - len(suffix)] + suffix
        i += 1
    taken.add(cleaned)
    return cleaned


def _style_header_row(ws) -> None:
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F5496")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _write_category_sheet(ws, records_for_cat: list[ImageRecord], link_map: dict[str, str]) -> int:
    """Write one category's rows into ws. Returns the number of painting rows."""
    ws.append(SHEET_HEADERS)
    _style_header_row(ws)

    by_cluster: dict[int, list[ImageRecord]] = defaultdict(list)
    for rec in records_for_cat:
        by_cluster[rec.cluster_id].append(rec)

    rows = []
    for group in by_cluster.values():
        primary = next((r for r in group if r.is_primary), group[0])
        mockups = [r for r in group if r is not primary]
        rows.append((primary, mockups))
    rows.sort(key=lambda x: x[0].filename.lower())

    for primary, _mockups in rows:
        primary_link = link_map.get(primary.filename.lower(), "")
        ws.append([
            primary_link,                          # link (Drive URL or empty)
            "",                                    # name (manual)
            "",                                    # short description (manual)
            "",                                    # long description (manual)
            "",                                    # technical description (manual)
        ])

    for i, w in enumerate(SHEET_COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = wrap
    ws.freeze_panes = "A2"
    return len(rows)


def write_xlsx(
    by_category: dict[str, list[ImageRecord]],
    output: Path,
    link_map: dict[str, str],
) -> tuple[int, int]:
    """Write one worksheet per category. Returns (categories, total_rows)."""
    wb = Workbook()
    # We don't know the first category name yet, so reuse the default sheet
    # for it instead of leaving an empty placeholder.
    categories = sorted(by_category.keys(), key=str.lower)
    if not categories:
        wb.active.title = "Paintings"
        wb.save(output)
        return 0, 0

    taken_sheet_names: set[str] = set()
    total = 0
    first_ws = wb.active
    for i, category in enumerate(categories):
        sheet_name = safe_sheet_name(category, taken_sheet_names)
        ws = first_ws if i == 0 else wb.create_sheet()
        ws.title = sheet_name
        total += _write_category_sheet(ws, by_category[category], link_map)

    wb.save(output)
    return len(categories), total


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input_dir", type=Path, help="Folder with category subfolders of paintings")
    ap.add_argument("-o", "--output", type=Path, default=Path("paintings.xlsx"))
    ap.add_argument("--link-map", type=Path, default=None,
                    help="CSV (filename,url) from Drive Apps Script export")
    ap.add_argument("--phash-threshold", type=int, default=6,
                    help="Hamming distance for pHash near-dup grouping (default 6)")
    ap.add_argument("--orb-min-matches", type=int, default=25,
                    help="Minimum good ORB matches to merge clusters as mockup-of-X (default 25)")
    ap.add_argument("--no-feature-match", action="store_true",
                    help="Skip ORB feature matching (mockups won't be grouped with their source)")
    args = ap.parse_args(argv)

    if not args.input_dir.is_dir():
        ap.error(f"{args.input_dir} is not a directory")

    use_orb = not args.no_feature_match and HAS_CV2
    if not args.no_feature_match and not HAS_CV2:
        print("! opencv-python not installed; mockup-vs-source matching disabled.", file=sys.stderr)
        print("  Install with: pip install opencv-python-headless numpy", file=sys.stderr)

    print(f"Scanning {args.input_dir} ...")
    records: list[ImageRecord] = []
    for path, category in iter_image_files(args.input_dir):
        rec = load_record(path, category, compute_orb=use_orb)
        if rec is not None:
            records.append(rec)
    print(f"  loaded {len(records)} images")

    if not records:
        print("No images found.", file=sys.stderr)
        return 1

    # Cluster within each category. A painting that lives in two folders
    # is two records (different paths) and gets its own row in each
    # category's sheet.
    by_category: dict[str, list[ImageRecord]] = defaultdict(list)
    for rec in records:
        by_category[rec.category].append(rec)

    cluster_id_offset = 0
    for category, recs in by_category.items():
        cluster_by_phash(recs, args.phash_threshold)
        if use_orb:
            merge_clusters_by_orb(recs, args.orb_min_matches)
        pick_primaries(recs)
        # Make cluster IDs unique across categories so downstream grouping
        # by cluster_id never collides between sheets.
        local_max = max((r.cluster_id for r in recs), default=-1)
        for r in recs:
            r.cluster_id += cluster_id_offset
        cluster_id_offset += local_max + 1

    link_map: dict[str, str] = {}
    if args.link_map:
        link_map = load_link_map(args.link_map)
        print(f"  loaded {len(link_map)} filename->URL mappings")

    n_categories, total_rows = write_xlsx(by_category, args.output, link_map)
    print(f"Wrote {args.output}: {n_categories} categories, {total_rows} painting rows")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
