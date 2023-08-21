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
function Renderer(reportName, range, queryCommand, columnTitle, isCheckIn, referenceDateTime) {
  this.reportName = reportName;
  this.range = range;
  this.queryCommand = queryCommand;
  this.columnTitle = columnTitle;
  this.columnCount =0;
  this.rowCount = 0;
  this.isCheckIn = isCheckIn;
  this.referenceDateTime = referenceDateTime;
}

/**
 * rendering 을 한다.
 * Table 한개를 만든다.
 */
Renderer.prototype.render = function() {
  var data = this.gather();
  var sb = new StringBuilder();
  sb.append("<table class='gmail-table' style='border: solid 2px #DDEEEE; border-collapse: collapse; border-spacing: 0; font: normal 14px Roboto sans-serif; margin: 10px 0 0 60px; width: 60%;'>");
  sb.append("<thead>");
  sb.append("<tr>");
  this.columnTitle.forEach((title, index) => {
    sb.append("<th class='");
    sb.append("col");
    sb.append(index);
    if(index > 0) {
      sb.append("' style='background-color: #DDEFEF; border: solid 1px black; color: #336B6B; padding: 4px; text-align: center; text-shadow: 1px 1px 1px #fff;'>");
    }
    else {
      sb.append("' style='display:none'>");
    }
    sb.append(title);
    sb.append("</th>");
  });
  sb.append("</thead>");
  sb.append("<tbody>");
  sb.append("</tr>");
  data.forEach(row => {
    sb.append("<tr>");
    // td 에 background-color 를 주기 위한 criteria 
    let colDate;
    row.forEach((col, index) => {
      sb.append("<td class='");
      sb.append("col");
      sb.append(index);
      if(index > 0) {
        sb.append("' style='border: solid 1px #DDEEEE; color: #333; padding: 4px; text-align: center; text-shadow: 1px 1px 1px #fff;");
        if(!this.referenceDateTime) {
          // console.log("SKIP");
        }
        else {
          if(new Date(this.referenceDateTime).getTime() > new Date(colDate).getTime()) {
            // row 에 background
            sb.append(" background-color: #ddd !important");
          }
        }
        sb.append("'>");     
      }
      else {
        sb.append("' style='display:none'>");
        colDate = col;
      }  
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
  // arrival 과 checkIn 은 항상 동일한 checkInListsSheet
  // checkOut 은 동일하거나 다르거나 checkOutListsSheet
  var sheetId = this.isCheckIn ? checkInListsSheet.getSheetId() : checkOutListsSheet.getSheetId();
  var url = ws.getUrl().replace("/edit", "");
  var request = url + '/gviz/tq?gid=' + sheetId + '&range=' + this.range + '&tq=' + encodeURIComponent(this.queryCommand);
  var request_result = UrlFetchApp.fetch(request).getContentText();
  // get json object
  var _from = request_result.indexOf("{");
  var _to   = request_result.lastIndexOf("}")+1;  
  var jsonText = request_result.slice(_from, _to);
  var parsedObject = JSON.parse(jsonText);
  if(parsedObject.status !== 'ok') {
    console.log("ERROR ", this.queryCommand, request);
    throw new Error(this.queryCommand + " : " + JSON.stringify(parsedObject));
  }
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
