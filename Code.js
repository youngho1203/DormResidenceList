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
//
function onEdit(e) {
  if(!e) {
    return;
  }
  const range_modified = e.range;
  if(range_modified.getColumn() !== CHECK_OUT_COLUMN ) {
    return;
  }
  // change style
  changeStyleForCheckOut(range_modified);
}

function changeStyleForCheckOut(range) {
  var _range = range.offset(0,0,1, (LAST_COLUMN - 4));
  var font_family = range.getFontFamily();
  var text_color = range.getValue() ? "#980000" : "black";
  var style_builder = _range.getTextStyle().copy().setUnderline(false).setForegroundColor(text_color).setFontFamily(font_family);
  _range.setTextStyle(style_builder.setStrikethrough(range.getValue()).build());
}
