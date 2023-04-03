const ws = SpreadsheetApp.getActiveSpreadsheet();
const listsSheet = ws.getSheetByName("현황");
const CONFIG_SHEET_NAME = 'Config';
const configSheet = ws.getSheetByName(CONFIG_SHEET_NAME);
// 생년월일 Column
const BIRTH_DAY_COLUMN = 7;
// CheckOut Column
const CHECK_OUT_COLUMN = 4;

/** 
 * Creates the menu items for user to run scripts on drop-down.
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
  listsSheet.getRange("D3:E" + last_low).getValues().forEach((value, index) => { if(!value[0]) {
    var address = listsSheet.getRange("Q" + (3 + index)).getValue();
    if(address !== '') {
      const studentName = listsSheet.getRange("F" + (3 + index)).getValue();
      const emailCellData = [listsSheet.getRange("Q" + (3 + index)).getBackground(), value[1], studentName, address ];
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

// Regular expression to check if string is valid date
const DATE_PATTERN = /^(\d{4})(-|\/|\. )(0?[1-9]|1[012])(-|\/|\. )(0?[1-9]|[12][0-9]|3[01])$/;

function onEdit(e) {
  const range_modified = e.range;
  if(range_modified.getColumn() === BIRTH_DAY_COLUMN) {
    // format check
    var dateValue = range_modified.getDisplayValue();
    if(!DATE_PATTERN.test(dateValue)) {
      range_modified.setValue(dateValue.replaceAll(/\. ?(\d)/g, ". $1"));
    }
    return;
  };
  if(range_modified.getColumn() !== CHECK_OUT_COLUMN ) {
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
  // has_extension column
  _extension_cell.setTextStyle(cell_text_style);
}

/**
 * 이 빠진 row 를 찾는다.
 */
function findEmptyRow(preDefinedArray) {
  lastRow = listsSheet.getLastRow();
  listsSheet.getRange("A3:E" + lastRow).getValues().forEach(values, index => {
    if(isCellEmpty(values[4])) {

    }
  });
}

function moveRow(fromRow, toRow) {

}

// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}