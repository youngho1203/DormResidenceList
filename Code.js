const ws = SpreadsheetApp.getActiveSpreadsheet();
const listsSheet = ws.getSheetByName("현황");
const configSheet = ws.getSheetByName("Config");
/** 
 * Creates the menu item "Select EmailAddress" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gather Data')
      .addItem('Select EmailAddress', 'selectEmailAddress')
      .addToUi();
}

function selectEmailAddress() { 
  var emailAddress = [];
  var last_low = listsSheet.getLastRow();
  listsSheet.getRange("E3:E" + last_low).getValues().forEach((studentId, index) => { if(studentId > 0) {
    var address = listsSheet.getRange("P" + (3 + index)).getValue();
    if(address !== '') {
      const studentName = listsSheet.getRange("F" + (3 + index)).getValue();
      const emailCellData = [listsSheet.getRange("P" + (3 + index)).getBackground(), studentId, studentName, address ];
      emailAddress.push(emailCellData);
    }
  }});
  addEmailAddressSheet(emailAddress);
}

const sheetName = "ActiveEmailAddress";

function addEmailAddressSheet(emailAddress) {
  // 
  var emailSheet = ws.getSheetByName(sheetName);
  if(!emailSheet) {
    ws.insertSheet(sheetName);
  }
  emailSheet = ws.getSheetByName(sheetName);
  var lastColumn = emailSheet.getLastColumn() + 1;
  lastColumn++;
  emailAddress.forEach((a, index) => emailSheet.getRange(1 + index, lastColumn).setValue(a[1]).setBackground(a[0]));
  lastColumn++;
  emailAddress.forEach((a, index) => emailSheet.getRange(1 + index, lastColumn).setValue(a[2]).setBackground(a[0]));
  lastColumn++;
  emailAddress.forEach((a, index) => emailSheet.getRange(1 + index, lastColumn).setValue(a[3]).setBackground(a[0]));  
}

function onEdit(e) {
  const range_modified = e.range;
  if(range_modified.getColumn() !== 4) {
    return;
  }
  var row = range_modified.getRow();
  // has extension column
  // cell text style 이 다른 column 과 다름.
  var _extension_cell = listsSheet.getRange(row, 14); 
  var cell_text_style = _extension_cell.getTextStyle();
  var lastColumn = listsSheet.getLastColumn();
  var _range = listsSheet.getRange(row, 4, 1, lastColumn);
  var font_family = range_modified.getFontFamily();
  var text_color = range_modified.getValue() ? "#980000" : "black";
  var style_builder = _range.getTextStyle().copy().setUnderline(false).setForegroundColor(text_color).setFontFamily(font_family);
  _range.setTextStyle(style_builder.setStrikethrough(range_modified.getValue()).build());
  // studentId column
  var id_range = listsSheet.getRange(row, 5);
  var studentId = -1 * id_range.getValue();
  id_range.setValue(studentId);
  // has_extension column
  _extension_cell.setTextStyle(cell_text_style);
}