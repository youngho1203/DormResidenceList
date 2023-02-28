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
function doPrintChangeRequest() {
    const surveySheet = SpreadsheetApp.openById(ARRIVAL_SURVEY_ID);
    var config = surveySheet.getSheetByName(CONFIG_SHEET_NAME);
  var last_low = listsSheet.getLastRow();
  // D : 퇴실 marker
  // Q : Email address
  listsSheet.getRange("D3:Q" + last_low).getValues().forEach((value, index) => { if(!value[0]) {
    //
    var studentInfo = getStudentInfo(value[1]);
    var dormitoryInfo = getDormitoryInfo(config, studentInfo.roomNumber);
    var data = buildData(studentInfo, dormitoryInfo);
    var pdf_url = doProcessInsurance(data);
  }});
}

function doProcessInsurance(data) {
  // Retreive the template file and destination folder.
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
  var pdfName = data.studentId + '_' + data.문서번호;
  var pdf = createPDF(pdfFolder, document.getId(), pdfName);
  template_copy.setTrashed(true);
  //
  return pdf.getUrl();
}
