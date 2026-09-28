/**
 * Sample (Bean & Leaf, fictional data): import and clean monthly CSV exports from a Drive folder.
 * NOT YET RUN in a live Google account — test before showing it as finished work.
 *
 * Setup: Extensions > Apps Script, paste this file, add the script property EXPORT_FOLDER_ID
 * (Project Settings > Script properties), reload the spreadsheet, then use the Automation menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automation')
    .addItem('Import & clean exports', 'importExports')
    .addToUi();
}

/** Reads every CSV in the export folder, drops blank and duplicate lines, writes one clean table. */
function importExports() {
  const folderId = PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID');
  if (!folderId) throw new Error('Set the EXPORT_FOLDER_ID script property first.');
  const book = SpreadsheetApp.getActive();
  const sheet = book.getSheetByName('Clean') || book.insertSheet('Clean');
  const files = DriveApp.getFolderById(folderId).getFilesByType(MimeType.CSV);
  const seen = new Set();
  const rows = [];
  let header = null;
  while (files.hasNext()) {
    const [head, ...lines] = Utilities.parseCsv(files.next().getBlob().getDataAsString('UTF-8'));
    if (!header) header = head;
    for (const line of lines) {
      const key = line.join('\u0001');
      if (!line[0] || seen.has(key)) continue; // blank or duplicate line
      seen.add(key);
      rows.push(line.map((v) => String(v).replace(/\s+/g, ' ').trim()));
    }
  }
  if (!header) {
    book.toast('No CSV files found in the export folder', 'Automation');
    return;
  }
  // setValues() needs a rectangle: pad short rows, trim long ones to one width.
  const width = rows.reduce((w, r) => Math.max(w, r.length), header.length);
  const pad = (r) => r.concat(Array(width - r.length).fill('')).slice(0, width);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, width).setValues([pad(header)]);
  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows.map(pad));
  book.toast(`${rows.length} clean rows imported`, 'Automation');
}
