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
function Renderer(reportName, range, queryCommand, columnTitle) {
  this.reportName = reportName;
  this.range = range;
  this.queryCommand = queryCommand;
  this.columnTitle = columnTitle;
  this.columnCount =0;
  this.rowCount = 0;
}

/**
 * rendering 을 한다.
 * Table 한개를 만든다.
 */
Renderer.prototype.render = function() {
  var data = this.gather();
  var sb = new StringBuilder();
  sb.append("<table class='gmail-table'>");
  sb.append("<thead>");
  sb.append("<tr>");
  this.columnTitle.forEach((title, index) => {
    sb.append("<th class='");
    sb.append("col");
    sb.append((index + 1));
    sb.append("'>")
    sb.append(title);
    sb.append("</th>");
  });
  sb.append("</thead>");
  sb.append("<tbody>");
  sb.append("</tr>");
  data.forEach(row => {
    sb.append("<tr>");
    row.forEach((col, index) => {
      sb.append("<td class='");
      sb.append("col");
      sb.append((index + 1));
      sb.append("'>");
      sb.append(col);
      sb.append("</td>");
    });
    sb.append("</tr>");
  });
  sb.append("</tbody>");
  sb.append("</table>");
  return sb;
}

/**
 * rendering 을 위한 data 를 만든다.
 */
Renderer.prototype.gather = function() {
  var sheetId =currentListsSheet.getSheetId();
  var url = ws.getUrl().replace("/edit", "");
  var request = url + '/gviz/tq?gid=' + sheetId + '&range=' + this.range + '&tq=' + encodeURIComponent(this.queryCommand);  
  var request_result = UrlFetchApp.fetch(request).getContentText();    
  // get json object
  var _from = request_result.indexOf("{");
  var _to   = request_result.lastIndexOf("}")+1;  
  var jsonText = request_result.slice(_from, _to);
  var parsedObject = JSON.parse(jsonText);
  this.columnCount = parsedObject.table.cols.length;
  var result = [];
  parsedObject.table.rows.forEach(row => {
    var rowValue = row.c;
    var _row = [];
    for(var k=0; k<this.columnCount; k++) {
      if(!rowValue[k]) {
        _row.push('');
      }
      else {
        _row.push(rowValue[k].v);
      }
    }
    result.push(_row);
  });
  this.rowCount = result.length;
  
  return result;
}
