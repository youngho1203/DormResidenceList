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
function Renderer(activeSheet, reportName, reportContent, range, queryCommand) {
  this.activeSheet = activeSheet;
  this.reportName = reportName;
  this.reportContent = reportContent;
  this.range = range;
  this.queryCommand = queryCommand;
  this.columnCount =0;
  this.rowCount = 0;
}

Renderer.prototype.render = function() {
  var data = this.gather();
  var sb = new StringBuilder();
  sb.append("<table border='1'>");
  data.forEach(row => {
    sb.append("<tr>");
    row.forEach(col => { 
      sb.append("<td>");
      sb.append(col);
      sb.append("</td>");
    });
    sb.append("</tr>");
  });
  sb.append("</table>");
  return sb;
}

Renderer.prototype.gather = function() {
  var fileId = ws.getId();
  var sheetId =currentListsSheet.getSheetId();

  // console.log("WS URL : ", ws.getUrl());
  // https://docs.google.com/spreadsheets/d/1rDZ2t9fJUX8iJZsjF2gGHSWSvl1_X42Ji89-gK4H9PU/edit
  var url = ws.getUrl().replace("/edit", "");
  // var request = 'https://docs.google.com/spreadsheets/d/' + fileId + '/gviz/tq?gid=' + sheetId + '&range=' + rangeA1 + '&tq=' + encodeURIComponent(sqlText);
  var request = url + '/gviz/tq?gid=' + sheetId + '&range=' + this.range + '&tq=' + encodeURIComponent(this.queryCommand);  
  var request_result = UrlFetchApp.fetch(request).getContentText();
  // console.log("Request Result >>>> ", request_result.length);     
  // get json object
  var from = request_result.indexOf("{");
  var to   = request_result.lastIndexOf("}")+1;  
  var jsonText = request_result.slice(from, to);
  var parsedObject = JSON.parse(jsonText);
  this.columnCount = parsedObject.table.cols.length;
  var result = [];
  parsedObject.table.rows.forEach(row => {
    var rowValue = row.c;
    var row = [];
    for(var k=0; k<this.columnCount; k++) {
      if(!rowValue[k]) {
        row.push('');
      }
      else {
        row.push(rowValue[k].v);
      }
    }
    result.push(row);
  });
  this.rowCount = result.length;
  
  return result;
}
