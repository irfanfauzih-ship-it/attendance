// Attendance backend. Paste into Extensions > Apps Script of your Google Sheet.
const FOLDER_NAME = 'Attendance photos';
const HEADERS = ['Recorded at', 'Device time', 'Name', 'Shift', 'Type', 'Latitude', 'Longitude', 'Accuracy (m)', 'Location', 'Map', 'Photo'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Attendance') || ss.insertSheet('Attendance');
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

    const now = new Date();
    const stamp = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss');
    const safeName = String(d.name).replace(/[^\w-]+/g, '_');
    const bytes = Utilities.base64Decode(d.photo.split(',')[1]);
    const file = folder.createFile(
      Utilities.newBlob(bytes, 'image/jpeg', stamp + '_' + safeName + '_' + String(d.type).replace(' ', '-') + '.jpg')
    );

    sheet.appendRow([
      now, d.deviceTime, d.name, d.shift, d.type, d.lat, d.lng, d.accuracy, d.address,
      'https://maps.google.com/?q=' + d.lat + ',' + d.lng, file.getUrl()
    ]);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return reply({ ok: true, message: 'Attendance endpoint is running' }); }

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
