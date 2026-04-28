/**
 * drive_link_export.gs
 *
 * Google Apps Script. Run once inside Google Drive to produce a
 * filename -> shareable-URL CSV that scan_paintings.py can join against.
 *
 * Setup:
 *   1. https://script.google.com  ->  New project
 *   2. Paste this whole file in.
 *   3. Replace ROOT_FOLDER_ID below with the ID of your paintings folder
 *      (the part after /folders/ in the Drive URL).
 *   4. Save, then run `exportPaintingLinks`. First run asks for permissions.
 *   5. A new Sheet "Painting links" is created in your Drive root with two
 *      columns: filename, url.
 *   6. File -> Download -> Comma-separated values (.csv) -> save as links.csv
 *   7. Pass it to the scanner:  python scan_paintings.py ./paintings --link-map links.csv
 *
 * What it does:
 *   - Recursively walks ROOT_FOLDER_ID.
 *   - Sets each image to "Anyone with the link can view" (idempotent).
 *   - Records filename + the standard shareable URL.
 *
 * Caveats:
 *   - If two files share a filename (e.g. duplicate uploads), the last one
 *     wins in the CSV. Rename or move duplicates beforehand if that matters.
 *   - For very large folders (10k+ files) this can take several minutes.
 */

const ROOT_FOLDER_ID = 'PUT_YOUR_DRIVE_FOLDER_ID_HERE';

const IMAGE_MIME_PREFIXES = ['image/'];

function exportPaintingLinks() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const sheet = SpreadsheetApp.create('Painting links').getActiveSheet();
  sheet.appendRow(['filename', 'url', 'category', 'mime']);

  const rows = [];
  walk_(root, '', rows);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('Wrote ' + rows.length + ' rows.');
}

function walk_(folder, categoryPath, rows) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const mime = f.getMimeType();
    if (!IMAGE_MIME_PREFIXES.some(p => mime.indexOf(p) === 0)) continue;
    try {
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // Some shared-drive items reject permission changes; keep going.
    }
    rows.push([f.getName(), f.getUrl(), categoryPath, mime]);
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    const sub_path = categoryPath ? categoryPath + '/' + sub.getName() : sub.getName();
    walk_(sub, sub_path, rows);
  }
}
