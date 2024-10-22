// Compiled using undefined undefined (TypeScript 4.9.5)
// Compiled using undefined undefined (TypeScript 4.9.5)
function deleteEmptyRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 0; i--) {
        if (data[i].every(function (cell) { return cell === ""; })) {
            sheet.deleteRow(i + 1);
        }
    }
}
function onOpen(e) {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem("サイドバーを表示する", "showSidebar")
        .addToUi();
}
function showSidebar() {
    var sidebarUi = HtmlService.createHtmlOutputFromFile("gas/sidebar").setTitle("title");
    SpreadsheetApp.getUi().showSidebar(sidebarUi);
}
