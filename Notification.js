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
const configSheet = ws.getSheetByName("Config");
const reportSheet = ws.getSheetByName("Report");
const historySheet = ws.getSheetByName("Report History");
const currentListsSheetName = configSheet.getRange("K2").getValue();
const currentListsSheet = ws.getSheetByName(currentListsSheetName);

/**
 * 주기적으로 check out report 를 email 로 보낸다.
 */
function sendNotification() {
  var lastLow = configSheet.getLastRow();
  configSheet.getRange("A15:G" + lastLow).getValues().forEach(array => {
    // 순번, reportName, report time, target email list, queryRange, queryCommand, templateName  
    var modifyValue = isModified(array[1]);
    if(modifyValue[0] != 0) {
      sendEmail(array[1], modifyValue, array[3], array[4], array[5], array[6]);
    }
  });
}

/**
 * 현황 List 가 update 되었는지 확인한다.
 * 0 : no change
 * 1 : checkIn change
 * 2 : checkOut change
 * 3 : checkIn, checkOut change
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
 */
function sendEmail(reportName, reportContent, targetEmailList, queryRange, queryCommand, templateName) {
  var data = [new Date(), reportName, '', ''];
  try{
    var subject = "[광토기숙사] " + reportName + '가 Update 되었습니다.';
    var templateFile_1 = HtmlService.createTemplateFromFile("인원 변동 일일 보고 앞부분");
    templateFile_1.date = data[0];
    templateFile_1.count = "?";
    //
    var message_1 = templateFile_1.evaluate().getContent(); 
    var templateFile_2 = HtmlService.createTemplateFromFile("인원 변동 일일 보고 뒷부분");
    templateFile_2.url = ws.getUrl();
    templateFile_2.gid = reportSheet.getSheetId();
    var message_2 = templateFile_2.evaluate().getContent();

    var message =    new Renderer(ws, reportName, reportContent, queryRange, queryCommand).render().toString();
    
    message = message_1 + message + message_2;
    // The code below will send an email with the current date and time.
    targetEmailList.split(',').forEach(address => {
      GmailApp.sendEmail(address, subject, '', { htmlBody: message });
    });
    data[2] = reportContent.slice(1).join('|');
    data[3] = 'SENT'

    // test_getSheetsQueryResult();
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
 * 변화가 있는지 확인하기 위한 문자열
 */
function getCurrentValue() {
  let lastLow = currentListsSheet.getLastRow();
  var checkIn = currentListsSheet.getRange("D3:E" + lastLow).getValues().filter(a => a[1] != '').map((a,index) => { return (a[0] ? index : a[0]) }).toString();  
  var checkOut = currentListsSheet.getRange("D3:E" + lastLow).getValues().filter(a => a[1] != '').map((a,index) => { return (a[0] ? a[0] : index) }).toString();  
  return hash(checkIn) + '|' + hash(checkOut);
}

/**
 * report 이전 상태 값을 구한다.
 */
function getLastValue(reportName) {
  var lastValue;
  let lastLow = historySheet.getLastRow();
  historySheet.getRange("B2:C" + lastLow).getValues().filter(array => array[0] == reportName).forEach(array => {
    lastValue = array[1];
  });
  return lastValue;
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