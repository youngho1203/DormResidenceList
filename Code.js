/**
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
const ws = SpreadsheetApp.getActiveSpreadsheet();
const listsSheet = ws.getActiveSheet();
const configSheet = ws.getSheetByName("Config");
// lastColumn
const LAST_COLUMN = listsSheet.getLastColumn();
// CheckOut Column 4 ('D')
const CHECK_OUT_COLUMN = configSheet.getRange("A12").getValue();
// Extension Column 15 ('O')
const EXTENSION_COLUMN = configSheet.getRange("B12").getValue();

function onEdit(e) {
  if(!e){
    return;
  }
  const range_modified = e.range;
  if(range_modified.getSheet().getName() != listsSheet.getName()) {
    return;
  }
  if(range_modified.getColumn() == 11){
    // 납부
    let value = range_modified.getValue() == 'o' ? _getNowDateISOFormattedString() : '';
    // 입사일 : R
    range_modified.offset(0,7,1,1).setValue(value);
    // V,W,X
    range_modified.offset(0,12,1,1).setValue(new Date().toString());
    return;
  }
  if(range_modified.getColumn() !== CHECK_OUT_COLUMN ) {
    return;
  }
  // 퇴사
  let value = range_modified.getValue() ? _getNowDateISOFormattedString() : '';
  // 퇴사일 : S
  range_modified.offset(0,15,1,1).setValue(value);
  // V,W,X
  value = range_modified.getValue() ? new Date().toString() : '';
  range_modified.offset(0,20,1,1).setValue(value);
  // change style
  changeStyleForCheckOut(range_modified);
}

function changeStyleForCheckOut(range) {
  // has extension column
  // cell text style 이 다른 column 과 다름.
  var _extension_cell = listsSheet.getRange(range.getRow(), EXTENSION_COLUMN); 
  var cell_text_style = _extension_cell.getTextStyle();
  var _range = range.offset(0,0,1, (LAST_COLUMN - 4));
  var font_family = range.getFontFamily();
  var text_color = range.getValue() ? "#980000" : "black";
  var style_builder = _range.getTextStyle().copy().setUnderline(false).setForegroundColor(text_color).setFontFamily(font_family);
  _range.setTextStyle(style_builder.setStrikethrough(range.getValue()).build());
  
  // has_extension column
  _extension_cell.setTextStyle(cell_text_style);
}
