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
const reportSheet = ws.getSheetByName("Daily Report");
const historySheet = ws.getSheetByName("Report History");
const currentListsSheetName = configSheet.getRange("K2").getValue();
const currentListsSheet = ws.getSheetByName(currentListsSheetName);

/**
 * 주기적으로 check out report 를 email 로 보낸다.
 */
function sendNotification() {
  var now = new Date();
  // simple trick to set Date
  reportSheet.getRange("A1").setValue(now.toISOString().substring(0, 10));
  var numberOfresidence = currentListsSheet.getRange("F1:J1").getValues()[0];
  var lastLow = configSheet.getLastRow();
  configSheet.getRange("A15:G" + lastLow).getValues().forEach(array => {
    // 순번, reportName, report time, target email list, queryRange, partialQueryCommand, templateName
    var modifyValue = isModified(array[1]);
    if(modifyValue[0] != 0) {
      sendEmail(now, array[1], modifyValue, array[3], array[4], array[5], array[6], numberOfresidence);
    }
  });
}

/**
 * 현황 List 가 update 되었는지 확인한다.
 * 0 : no change
 * 1 : checkIn change
 * 2 : checkOut change
 * 3 : checkIn, checkOut both change
 */
function isModified(reportName) {
  var lastValue = getLastValue(reportName).split('|');
  var currentValue = getCurrentValue().split('|');

  if(lastValue[0] == currentValue[0] && lastValue[1] == currentValue[1]) {
    return [0, currentValue[0], currentValue[1]];
  }
  else if(lastValue[0] != currentValue[0] && lastValue[1] == currentValue[1]) {
    return [1, currentValue[0], currentValue[1]];
  }
  else if(lastValue[0] == currentValue[0] && lastValue[1] != currentValue[1]) {
    return [2, currentValue[0], currentValue[1]];
  }
  else {
    return [3, currentValue[0], currentValue[1]];
  }
}

/**
 * email 을 보낸다.
 * @param partialQueryCommand 는 queryCommand 의 앞부분만 가지고 있다.
 */
function sendEmail(now, reportName, reportContent, targetEmailList, queryRange, partialQueryCommand, templateName, numberOfresidence) {
  var data = [now, reportName, '', ''];
  try{
    //
    var templateFile_1 = HtmlService.createTemplateFromFile(templateName + " 앞부분");
    templateFile_1.date = data[0];
    templateFile_1.numberOfresidence = numberOfresidence;
    //
    var templateFile_2 = HtmlService.createTemplateFromFile(templateName + " 뒷부분");
    templateFile_2.url = ws.getUrl();
    templateFile_2.gid = reportSheet.getSheetId();
    //
    var htmlMessage = new StringBuilder();
    htmlMessage.append(templateFile_1.evaluate().getContent());
    
    var title = getTitle(partialQueryCommand);
    var checkInQueryCommand = partialQueryCommand + " D=False AND R = date '" + data[0].toISOString().substring(0, 10) + "'";
    var checkOutQueryCommand = partialQueryCommand + " D=True AND S = date '" + data[0].toISOString().substring(0, 10) + "'";
    //
    if(reportContent[0] == 1) {
      // checkIn Only
      _doRender(htmlMessage, reportName, queryRange, checkInQueryCommand, title, "신규 입사생 수");
    }
    else if(reportContent[0] == 2) {
      // checkOut Only  
      _doRender(htmlMessage, reportName, queryRange, checkOutQueryCommand, title, "신규 퇴사생 수");  
    }
    else {
      // checkIn, CheckOut both
      _doRender(htmlMessage, reportName, queryRange, checkInQueryCommand, title, "신규 입사생 수");
      _doRender(htmlMessage, reportName, queryRange, checkOutQueryCommand, title, "신규 퇴사생 수");
    }
    //
    htmlMessage.append(templateFile_2.evaluate().getContent());
    //
    var subject = "[광토기숙사] " + reportName + '가 Update 되었습니다.';
    targetEmailList.split(',').forEach(address => {
      GmailApp.sendEmail(address, subject, '', { htmlBody: htmlMessage.toString() });
    });
    data[2] = reportContent.slice(1).join('|');
    data[3] = 'SENT'
  }
  catch(ex) {
    console.log(ex);
    data[2] = '0|0';
    data[3] = ex;
  }
  // 
  historySheet.appendRow(data);
}

/**
 *
 */
function _doRender(htmlMessage, reportName, queryRange, queryCommand, title, reportTitle) {
  var renderer = new Renderer(reportName, queryRange, queryCommand, title); 
  var checkInMessage = renderer.render();
  htmlMessage.append("<div class='sub-title' style='font: normal 14px Roboto, sans-serif; margin: 10px 0 6px 0;'>");
  htmlMessage.append("• ");
  htmlMessage.append(reportTitle);
  htmlMessage.append(" : [ ");
  htmlMessage.append(renderer.rowCount);
  htmlMessage.append(" ]</div>");
  htmlMessage.append(checkInMessage);
}

/**
 * 변화가 있는지 확인하기 위한 문자열
 */
function getCurrentValue() {
  let lastLow = currentListsSheet.getLastRow();
  var checkIn = currentListsSheet.getRange("E3:R" + lastLow).getValues().filter(a => a[0] != '').filter(a => a[13] != '').map(a => { return a[13].toISOString().substring(0,10) }).toString();  
  var checkOut = currentListsSheet.getRange("E3:S" + lastLow).getValues().filter(a => a[0] != '').filter(a => a[14] != '').map(a => { return a[14].toISOString().substring(0,10) }).toString();
  return hash(checkIn) + '|' + hash(checkOut);
}

/**
 * report 이전 상태 값을 구한다.
 */
function getLastValue(reportName) {
  var lastValue ='0|0';
  let lastLow = historySheet.getLastRow();
  historySheet.getRange("B2:C" + lastLow).getValues().filter(array => array[0] == reportName).forEach(array => {
    lastValue = array[1];
  });
  return lastValue;
}

function getTitle(partialQueryCommand) {
  // SELECT xxxx WHERE statement
  let cols = partialQueryCommand.substring(7, partialQueryCommand.indexOf("WHERE")).split(",");
  // 2열 이 제목이다.
  let rangeList = cols.map(c => (c.trim() + 2));
  return currentListsSheet.getRangeList(rangeList).getRanges().map(r => r.getValue());
}

/**
 * Simple string hash for checking two string difference
 */
function hash(str) {
  var hash = 0,
  i, chr;
  if (str.length === 0) return hash;
  for (i = 0; i < str.length; i++) {
    chr = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + chr;
    hash |= 0; // Convert to 32bit integer
  }
  return hash;
}