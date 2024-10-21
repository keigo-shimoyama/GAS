function getParentFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents();
  return parentFolder.next();
}

function getFileList(folder, spreadsheetName) {
  const files = folder.getFiles();
  let fileList = [];
  const now = new Date();
  while (files.hasNext()) {
    const file = files.next();
    const createdDate = file.getDateCreated();
    const timeDifference = (now - createdDate) / (1000 * 60); // difference in minutes
    if (file.getName() !== spreadsheetName && timeDifference <= 5) {
      fileList.push({
        id: file.getId(),
        name: file.getName(),
      });
    }
  }
  return fileList.sort((a, b) => a.name.localeCompare(b.name));
}

function insertImages(sheet, fileList) {
  let row = 2;
  const column = 2; // B column
  fileList.forEach((file) => {
    const fileUrl = "https://lh3.google.com/u/0/d/" + file.id;
    const formula = `=IMAGE("${fileUrl}", 1)`;
    sheet.getRange(row, column).setFormula(formula);
    row++;
  });
}

function insertTimestamps(sheet, fileList) {
  let row = 2;
  const column = 1; // A column
  fileList.forEach((file) => {
    const timeFormatted = file.name
      .match(/_(\d{2})\.(\d{2})\./)
      .slice(1, 3)
      .join(":");
    sheet.getRange(row, column).setValue(timeFormatted);
    row++;
  });
}

function renameSpreadsheet(sheet, fileList) {
  const now = new Date();
  const formattedDate = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_"
  );
  const newSpreadsheetName =
    formattedDate + fileList[0].name.match(/[^\.]+/)[0] + "_";
  sheet.setName(newSpreadsheetName);
}

function insertImagesAndTimestamps() {
  const folder = getParentFolder();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const spreadsheetName = spreadsheet.getName();
  const fileList = getFileList(folder, spreadsheetName);
  if (fileList.length === 0) {
    Browser.msgBox("No new images found in the folder.");
    return;
  }
  insertImages(sheet, fileList);
  insertTimestamps(sheet, fileList);
  renameSpreadsheet(sheet, fileList);
}
