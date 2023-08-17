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
// checkIn, checkOut sheet split
const checkInSheetName = configSheet.getRange("J2").getValue();
const checkInListsSheet = ws.getSheetByName(checkInSheetName);
const checkOutSheetName = configSheet.getRange("K2").getValue();
const checkOutListsSheet = ws.getSheetByName(checkOutSheetName);
const reportTitleArray = ["신규 도착생 수", "신규 입사생 수", "신규 퇴사생 수"];

/**
 * 주기적으로 check out report 를 email 로 보낸다.
 */
function sendNotification() {
  var now = new Date();
  // simple trick to set Date
  reportSheet.getRange("A1").setValue(now);
  // @todo checkIn, checkOut sheet 를 분리함에 따라서, 보다 복잡해졌다.
  var numberOfresidence = getNumberOfCurrentResident();
  var lastLow = configSheet.getLastRow();
  configSheet.getRange("A15:G" + lastLow).getValues().forEach(array => {
    // 순번, reportName, report time, target email list, queryRange, partialQueryCommand, templateName
    var modifyValue = isModified(array[1]);
    if(modifyValue[0] > 0) {
      sendEmail(now, array[1], modifyValue, array[3], array[4], array[5], array[6], numberOfresidence);
    }
  });
}

/**
 * CheckIn, CheckOut sheet 를 분리함에 따라서 복잡해졌다. 
 */
function getNumberOfCurrentResident() {
  // @todo checkIn, checkOut sheet 분리 적용 implement
  return checkInListsSheet.getRange("F1:J1").getValues()[0];
}

/** 
 * 현황 List 가 update 되었는지 확인한다.
 * arrival | checkIn | checkOut
 * 000 : 0 : no change
 * 100 : 4 : arrival change only
 * 010 : checkIn change only
 * 001 : checkOut change only
 * 110 : arrival, checkIn both change
 * 101 : arrival, checkOut both change
 * 011 : checkIn, checkOut both change
 * 111 : arrival, checkIn, checkOut all change
 */
function isModified(reportName) {
  //
  let lastValue = getLastValue(reportName).split('|');
  let currentValue = getCurrentValue().split('|');
  let compareArray = [ 
    lastValue[0] == currentValue[0] ? 0 : 1,
    lastValue[1] == currentValue[1] ? 0 : 1, 
    lastValue[2] == currentValue[2] ? 0 : 1
  ];
  //
  let n = binArraytoInt(compareArray);

  return [n, ...currentValue];
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
    //
    var dateString = _getISOTimeZoneCorrectedDateString(data[0]);    
    var title = getTitle(partialQueryCommand);
    let arrivalQueryCommand = partialQueryCommand + " D=False AND U = date '" + dateString + "'";
    let checkInQueryCommand = partialQueryCommand + " D=False AND R = date '" + dateString + "'";
    let checkOutQueryCommand = partialQueryCommand + " D=True AND S = date '" + dateString + "'";
    //
    let n = reportContent[0];
    var updateCount = 0;
    if(n < 2) {
      // 001 : checkOut change only
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkOutQueryCommand, title, -1);
    }
    else if(n < 3) {
      // 010 : checkIn change only
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkInQueryCommand, title, 1);
    }
    else if (n < 4) {
      // 011 : checkIn, checkOut both change
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkInQueryCommand, title, 1);
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkOutQueryCommand, title, -1);
    }
    else if (n < 5) {
      // 100 : arrival change only
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, arrivalQueryCommand, title, 0);
    }
    else if(n < 6) {
      // 101 : arrival, checkOut both change
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, arrivalQueryCommand, title, 0);
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkOutQueryCommand, title, -1);
    }
    else if(n < 7) {
      // 110 : arrival, checkIn both change 
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, arrivalQueryCommand, title, 0);
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkInQueryCommand, title, 1);
    }
    else if(n < 8){
      // 111 : arrival, checkIn, checkOut all change
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, arrivalQueryCommand, title, 0);
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkInQueryCommand, title, 1);
      updateCount = updateCount + _doRender(htmlMessage, reportName, queryRange, checkOutQueryCommand, title, -1);
    }
    else {
      throw new Error("Not Supported [" + n + "]");
    }

    if(updateCount > 0) {
      //
      htmlMessage.append(templateFile_2.evaluate().getContent());
      //
      var subject = "DEV [광토기숙사] " + reportName + '가 Update 되었습니다.';
      targetEmailList.split(',').forEach(address => {
        GmailApp.sendEmail(address, subject, '', { htmlBody: htmlMessage.toString() });
      });
      data[2] = reportContent.slice(1).join('|');
      data[3] = 'SENT'
    }
    else {
      data[2] = reportContent.slice(1).join('|');
      data[3] = 'SKIP'      
    }
  }
  catch(ex) {
    console.log(ex);
    data[2] = '0|0|0';
    data[3] = ex.stack;
  }
  // 
  historySheet.appendRow(data);
}

/**
 * @return renderer.rowCount
 */
function _doRender(htmlMessage, reportName, queryRange, queryCommand, title, type) {
  /**
   * type 0 : arrival
   * type 1 : checkIn
   * type -1 : checkOut
   */
  let reportTitle = reportTitleArray[type];
  let isCheckIn = type > -1; 
  var renderer = new Renderer(reportName, queryRange, queryCommand, title, isCheckIn); 
  var message = renderer.render();
  htmlMessage.append("<div class='sub-title' style='font: normal 14px Roboto, sans-serif; margin: 10px 0 6px 0;'>");
  htmlMessage.append("• ");
  htmlMessage.append(reportTitle);
  htmlMessage.append(" : [ ");
  htmlMessage.append(renderer.rowCount);
  htmlMessage.append(" ]</div>");
  htmlMessage.append(message);
  //
  return renderer.rowCount;
}

/**
 * 변화가 있는지 확인하기 위한 문자열
 */
function getCurrentValue() {
  let lastLow = checkInListsSheet.getLastRow();
  //
  let arrival = checkInListsSheet.getRange("E3:U" + lastLow).getValues().filter(a => a[0] != '').filter(a => a[16] != '').map(a => { return _getISOTimeZoneCorrectedDateString(a[16])}).toString();
  //
  let checkIn = checkInListsSheet.getRange("E3:R" + lastLow).getValues().filter(a => a[0] != '').filter(a => a[13] != '').map(a => { return _getISOTimeZoneCorrectedDateString(a[13])}).toString();
  //
  lastLow = checkOutListsSheet.getLastRow();
  let checkOut = checkOutListsSheet.getRange("E3:S" + lastLow).getValues().filter(a => a[0] != '').filter(a => a[14] != '').map(a => { return _getISOTimeZoneCorrectedDateString(a[14])}).toString();
  return hash(arrival) + '|' + hash(checkIn) + '|' + hash(checkOut);
}

/**
 * report 이전 상태 값을 구한다.
 */
function getLastValue(reportName) {
  var lastValue ='0|0|0';
  let lastLow = historySheet.getLastRow();
  historySheet.getRange("B2:C" + lastLow).getValues().filter(array => array[0] == reportName).forEach(array => {
    lastValue = array[1];
  });
  return lastValue;
}

/**
 * table th 이름을 가지고 온다.
 * CheckIn, CheckOut sheet 가 동일하여야 한다.
 */
function getTitle(partialQueryCommand) {
  // SELECT xxxx WHERE statement
  let cols = partialQueryCommand.substring(7, partialQueryCommand.indexOf("WHERE")).split(",");
  // 2열 이 제목이다.
  let rangeList = cols.map(c => (c.trim() + 2));
  return checkInListsSheet.getRangeList(rangeList).getRanges().map(r => r.getValue());
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