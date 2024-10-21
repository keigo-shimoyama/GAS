function getParentFolder() {
    // 自身のスプレッドシートのIDを取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssId = ss.getId();

    // 親フォルダ(ファイル自身が格納されているフォルダ)を取得
    var parentFolder = DriveApp.getFileById(ssId).getParents();
    var folder = parentFolder.next();
    
    return folder;
}

function insertImages() {

    const folder = getParentFolder();
    const files = folder.getFiles();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Collect all files and their names
    let row = 2;
    let column = 2; // B column

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetName = spreadsheet.getName();

    let fileList = [];
    while (files.hasNext()) {
        const file = files.next();
        if (file.getName() !== spreadsheetName) { // スプレッドシート自身のファイル名を除外
            fileList.push({
                id: file.getId(),
                name: file.getName()
            });
        }
    }

    // Sort files by their names in ascending order
    fileList.sort((a, b) => a.name.localeCompare(b.name));

    // Insert sorted images into the spreadsheet
    fileList.forEach(file => {
        const fileUrl = 'https://lh3.google.com/u/0/d/' + file.id;
        const formula = '=IMAGE("' + fileUrl + '", 1)';
        sheet.getRange(row, column).setFormula(formula);
        row++;
    });

    // Insert timestamps into the spreadsheet
    row = 2;
    column = 1; // A column

    fileList.forEach(file => {
        const timeFormatted = file.name.match(/_(\d{2})\.(\d{2})\./).slice(1, 3).join(":");
        sheet.getRange(row, column).setValue(timeFormatted);
        row++;
    });
}
