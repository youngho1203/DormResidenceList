const ws = SpreadsheetApp.getActiveSpreadsheet();
const listsSheet = ws.getSheetByName("현황");
// 생년월일 Column
const BIRTH_DAY_COLUMN = 7;
// CheckOut Column
const CHECK_OUT_COLUMN = 4;

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

// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}