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
// 변경 신청서 docs Template File ( 적절하게 변경해야 함. ) 
const INSURANCE_TEMPLATE_FILE_ID = '1IizRaNHtEN7zCJTtyLy79gHdEiABeBbNl-vQV1YWOSw';
const INSURANCE_OUTPUT_FOLDER_NAME = 'HealthInsurancePdfFolder';

function showHealthInsuranceDialog() {
  // Display a modal dialog box with custom HtmlService content.
  var dialog = HtmlService.createHtmlOutputFromFile("HealthDialog.html");
  SpreadsheetApp.getUi().showModalDialog(dialog, '신청날짜를 입력하세요');
}

function getHealthDataFromFormSubmit(form) {
  if(form.issueDate == undefined || form.issueDate == '') {
    SpreadsheetApp.getUi().alert('신청날짜를 입력하세요.');
    return;
  }  
  //
  var issueDate = form.issueDate;  
  
  // DormitoryInfo 때문에 SurveySheet 가 필요하다.
  const surveySheet = SpreadsheetApp.openById(ARRIVAL_SURVEY_ID);
  var config = surveySheet.getSheetByName(CONFIG_SHEET_NAME);
  var last_low = listsSheet.getLastRow();
  // D : 퇴실 marker
  // Q : Email address
  listsSheet.getRange("D3:Q" + last_low).getValues().forEach((value, index) => { if(!value[0]) {
    //
    var results;
    try {
      var studentInfo = getStudentInfo(value[1]);
      var dormitoryInfo = getDormitoryInfo(config, studentInfo.roomNumber);
      var data = buildData(studentInfo, dormitoryInfo, issueDate);
      results = doProcessInsurance(data);
    }catch(e) {
      results = e;
    }    
    listsSheet.getRange("U" + (3+index)).setValue(results);
  }});
}

function doProcessInsurance(data) {
  // Retreive the template file and destination folder.
  console.log(data);
  var template_file = DriveApp.getFileById(INSURANCE_TEMPLATE_FILE_ID);
  var template_copy = template_file.makeCopy(template_file.getName() + "(Copy)");
  var document = DocumentApp.openById(template_copy.getId());
  //
  populateTemplate(document, data);
  //
  document.saveAndClose();
  // console.log(document.getId());
  // Cleans up and creates PDF.
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf  
  //
  // pdf file saved folder
  const pdfFolder = getFolderByName_(template_file, INSURANCE_OUTPUT_FOLDER_NAME);  
  // save pdf file
  // save file name pattern : studentId_문서번호
  var pdfName = data.StudentId + '_HealthInsurance';
  var pdf = createPDF(pdfFolder, document.getId(), pdfName);
  template_copy.setTrashed(true);
  //
  return pdf.getUrl();
}
