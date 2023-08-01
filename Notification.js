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
const historySheet = ws.getSheetByName("Report History");
const currentListsSheetName = configSheet.getRange("K2").getValue();
const currentListsSheet = ws.getSheetByName(currentListsSheetName);

/**
 * 주기적으로 check out report 를 email 로 보낸다.
 */
function sendNotification() {
  var lastLow = configSheet.getLastRow();
  configSheet.getRange("A15:F" + lastLow).getValues().forEach(array => {
    // 순번, reportName, report time, target email list, query command, templateName  
    var count = modifiedCount(array[1]);
    if(count > 0) {
      sendEmail(array[1], count, array[3], array[4], array[5]);
    }
  });
}

/**
 * 현황 List 가 update 되었는지 확인한다.
 */
function modifiedCount(reportName) {
  // @todo use query command
  var lastCount = getLastCount(reportName);
  let count = query(reportName);
  return (count - lastCount > 0 ? count : 0 );
}

/**
 * email 을 보낸다.
 */
function sendEmail(reportName, report, targetEmailList, queryCommand, templateName) {
  //
  var data = [new Date(), reportName, report, ''];
  try{
    var subject = "TEST : [광토기숙사] " + reportName + '가 Update 되었습니다.';
    var templateFile = HtmlService.createTemplateFromFile(templateName);
    templateFile.date = data[0];
    var message = templateFile.evaluate().getContent();    
    // The code below will send an email with the current date and time.
    targetEmailList.split(',').forEach(address => {
      GmailApp.sendEmail(address, subject, '', { htmlBody: message });
    });
    data[3] = 'SENT'
  }
  catch(ex) {
    data[2] = -1;
    data[3] = ex;
  }
  // 
  historySheet.appendRow(data);
}

/**
 * Query Commend 를 실행한다.
 */
function query(reportName) {
  // @todo query command 로 구현 필요.
  //
  // 퇴사학생 수를 구한다.
  let lastLow = currentListsSheet.getLastRow();
  return currentListsSheet.getRange("D3:E" + lastLow).getValues().filter(a => a[0] == true && a[1] != '').length;
}

/**
 * report 수를 구한다. ( 주의 : 이전 대비 증가한 값만 구한다.)
 */
function getLastCount(reportName) {
  var lastCount = 0;
  let lastLow = historySheet.getLastRow();
  historySheet.getRange("B2:C" + lastLow).getValues().filter(array => array[0] == reportName).forEach(array => {
    console.log("비교 ", lastCount, array[1]);
    if(array[1]> lastCount){
      lastCount = array[1];
    }
  });
  return lastCount;
}