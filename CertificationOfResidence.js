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
const APP_TITLE = 'CertificationOfResidence';
// 거주증명서 docs Template File ( 적절하게 변경해야 함. ) 
const TEMPLATE_FILE_ID = '1GfxiSCucUEGVgffaahLzetvKUIah-M1-XmA4cupd988'; // @TODO need to change
// ArrivalSurvey(응답) SpreadSheet File ( 적절하게 변경해야 함. ) 
const ARRIVAL_SURVEY_ID = '1ZliHOc0nihMy9l6SnLsTpnZ2aA5yAffg41uVN3qbMF8'; // 2023 광토 기숙사 ArrialSurvey Sheet List
// 생성된 거주 증명서 저장 Folder Name
const OUTPUT_FOLDER_NAME = 'CertificationOfResitancePdfFolder';
// 거주 증명서 발급된 Row Background Color 
const BUILD_ROW_BACKGROUND_COLOR = '#e0e0e0';
// Alphabet
const ALPHABET = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',  'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' ];

function showSelectDialog() {
  // Display a modal dialog box with custom HtmlService content.
  var dialog = HtmlService.createHtmlOutputFromFile("ManualDialog.html");
  SpreadsheetApp.getUi().showModalDialog(dialog, '발행할 학번, 발행번호, 발행날짜를 입력하세요');
}
function showAllDialog() {
  // Display a modal dialog box with custom HtmlService content.
  var dialog = HtmlService.createHtmlOutputFromFile("AutoDialog.html");
  SpreadsheetApp.getUi().showModalDialog(dialog, '발행할 첫번째 발행번호, 발행날짜를 입력하세요');
}
function getDataFromFormSubmit(form) {
  if(form.issuedNumber == undefined || form.issuedNumber == '') {
    SpreadsheetApp.getUi().alert('IssueNumber 를 입력하세요.');
    return;
  }  
  //
  var issuedNumber = form.issuedNumber;
  //      
  var matched = issuedNumber.match(/\d+$/);
  var namePart = issuedNumber.substring(0, matched.index);
  var serialPart = matched[0];  
  //
  backgroundReset();
  //
  if(form.studentId) {
    buildCertificationOfResidenceBySelect(form.studentId, namePart, serialPart, form.issueDate);
  }
  else {
    buildCertificationOfResidence(namePart, serialPart, form.issueDate);
  }
}

function backgroundReset() {
  const lastRow = listsSheet.getLastRow();
  const lastColumn = listsSheet.getLastColumn();
  listsSheet.getRange("C3:" + ALPHABET[lastColumn -1] + lastRow).setBackground("white");
}

function buildCertificationOfResidence(namePart, serialPart, issueDate) {
  //
  // gather all 'submitted medical report' and 'Not yet publish Certification of Residence'
  // '학번'부터 14 개 column
  const lastLow = listsSheet.getLastRow();
  listsSheet.getRange("D3:T" + lastLow).getValues().forEach((value, index) => {
    if(value[0]) {
      // 퇴사자
      var range = listsSheet.getRange("A" + (3+index));
      revokeBackground(range);
    }
    else {
      // console.log('buildCertificationOfResidence', value);
      // 기숙사비를 내고, 건강진단서를 제출한 학생들 중
      if(value[7] === 'o' && value[8] === 'o') {
        // 기숙사비를 내고, 건강진단서를 제출한 학생들 중
        if(isCellEmpty(value[16]) || !value[16].startsWith("https://drive.google.com/")) {
          // 아직 발급이 되지 않은 학생들 
          serialPart = buildCertificationOfResidenceBySelect(value[1], namePart, serialPart, issueDate);
        }
        else {
          console.log('SKIP', value[16] );
          var range = listsSheet.getRange("A" + (3+index));
          revokeBackground(range);
        }
      }
      else {
        console.log('SKIP', value[7], value[8] );
        var range = listsSheet.getRange("A" + (3+index));
        revokeBackground(range);     
      }
    }
  })
}

function revokeBackground(range) {
  var background_color = range.getBackground();
  var row = range.getRow();
  listsSheet.getRange("C" + row + ":T" + row).setBackground(background_color);   
}

function buildCertificationOfResidenceBySelect(studentId, namePart, serialPart, issueDate) {
  // left padding
  var paddedSerialPart = serialPart.toString().padStart(3,0);
  var urlOrError;
  try {
    var studentInfo = getStudentInfo(studentId);
    /**
     * 입사일은 따로 설정하지 않고 발행일과 동일하게 설정하면 되겠습니다. 
     * 규정 상 입사일로부터 14일 이내에 거주 신고를 하여야 하는데, 단체 접수 특성 상 이 기간이 맞지 않습니다. 
     * 또 여러 비슷한 이유로 일어나는 문제 발생을 막기 위하여 입사일=발행일로 통일하고 있습니다. 
     * - 재훈사감 -
     */
    studentInfo.checkInDate = issueDate;
    console.log('studentInfo', studentInfo);
    //
    const surveySheet = SpreadsheetApp.openById(ARRIVAL_SURVEY_ID);
    var config = surveySheet.getSheetByName(CONFIG_SHEET_NAME);
    var dormitoryInfo = getDormitoryInfo(config, studentInfo.roomNumber);
    console.log('dormitoryInfo', dormitoryInfo);
    if(dormitoryInfo === undefined) {
      throw new Error(studentInfo.roomNumber + "호의 기숙사 정보를 확인할 수 없습니다.");
    }
    var data = buildData(studentInfo, dormitoryInfo, issueDate);
    data.문서번호 = namePart + paddedSerialPart;
    /**
     * {
      'studentId': studentInfo.studentId,
      '문서번호': namePart + paddedSerialPart,
      '입주일' : studentInfo.checkInDate,
      '주소' : dormitoryInfo.주소 + ' ' + studentInfo.roomNumber + '호',
      'Address' : '#' + studentInfo.roomNumber + ', ' + dormitoryInfo.Address,
      'MoveInDate': studentInfo.checkInDate,
      'Name' : studentInfo.name,
      'Email': studentInfo.email,
      'Phone': studentInfo.phone, 
      'StudentIDNumber':studentId,
      'BirthDay':studentInfo.birthDay,
      // @todo : 날짜 지정 ????
      '발급일자': issueDate, // new Date().toISOString().substring(0, 10),
      '신청일자': issueDate
    };
    */
    //  
    urlOrError = doProcess(data);
    serialPart++
  }
  catch(e) {
    urlOrError = e;
  }
  //
  const lastRow = listsSheet.getLastRow(); 
  listsSheet.getRange("E3:E" + lastRow).getValues().forEach((id, index) => {
    if(id == studentId) {
      listsSheet.getRange("T" + (index + 3)).setValue(urlOrError)
      listsSheet.getRange("C" + (index + 3) + ":T" + (index + 3)).setBackground(BUILD_ROW_BACKGROUND_COLOR);
    }
  });
  return serialPart;
}

function buildData(studentInfo, dormitoryInfo, issueDate) {
    return {
      'StudentId': studentInfo.studentId,
      /* '문서번호': namePart + paddedSerialPart, */
      '입주일' : studentInfo.checkInDate,
      '주소' : dormitoryInfo.주소 + ' ' + studentInfo.roomNumber + '호',
      'Address' : '#' + studentInfo.roomNumber + ', ' + dormitoryInfo.Address,
      'MoveInDate': studentInfo.checkInDate,
      'Name' : studentInfo.name,
      'Email': studentInfo.email,
      'Phone': studentInfo.phone, 
      /* 'StudentIDNumber':studentId,*/
      'BirthDay':studentInfo.birthDay,
      // @todo : 날짜 지정 ????
      '발급일자': issueDate, // new Date().toISOString().substring(0, 10),
      '신청일자': issueDate
    };  
}
/**
 * DataSheet 에서 matching 되는 학생 정보를 찾는다. 
 */
function getStudentInfo(studentId) {
  // 입사 학생 총 수
  const numberOfData = listsSheet.getLastRow();   
  var studentData; 
  listsSheet.getRange("E3:E" + numberOfData).getValues().forEach((id, index) => {
    if(id == studentId) {
      studentData = listsSheet.getRange(index + 3, 1, 1, 17).getValues()[0];
    }
  });

  if(studentData){
    // 생년월일 Column
    if(isCellEmpty(studentData[8])){
      throw new Error("생년월일의 값이 설정되어 있어야 합니다.");
    };
    // 납부 Column
    if(studentData[11] !== 'o') {
      throw new Error("아직 기숙사비를 내지 않았습니다.");
    };

    console.log('studentData', studentData);
    return { 
      'studentId':studentData[4], 
      'name':studentData[5], 
      'birthDay':studentData[(BIRTH_DAY_COLUMN -1)].toISOString().substring(0, 10), // cell format 이 date 로 설정되어 있어야 한다.
      'checkInDate': '',
      'roomNumber':studentData[1],
      'phone': studentData[15],
      'email':studentData[16]
      };
  }
  return undefined;
}

function getDormitoryInfo(dormitoryConfigSheet, roomNumber) {
  var dormitoryData;
  var lastLow = dormitoryConfigSheet.getLastRow();
  dormitoryConfigSheet.getRange("B2:B" + lastLow).getValues().forEach((id, index) => {
    if(id == roomNumber) {
      dormitoryData = dormitoryConfigSheet.getRange(index + 2, 1, 1, 6).getValues()[0];
    }
  });
  console.log('dormitoryData', dormitoryData);
  if(dormitoryData){
    if(dormitoryData[4] == '' || dormitoryData[5] == '') {
      throw new Error(roomNumber + "의 주소를 확인할 수 없습니다.");
    }    
    return { 
      'noomNumber':dormitoryData[1],
      '주소':dormitoryData[4],
      'Address':dormitoryData[5] 
      };
  }
  return undefined;  
}

function doProcess(data) {  
  // Retreive the template file and destination folder.
  var template_file = DriveApp.getFileById(TEMPLATE_FILE_ID);
  var template_copy = template_file.makeCopy(template_file.getName() + "(Copy)");
  var document = DocumentApp.openById(template_copy.getId());
  //
  populateTemplate(document, data);
  //
  document.saveAndClose();
  // Cleans up and creates PDF.
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf  
  // pdf file saved folder
  const pdfFolder = getFolderByName_(template_file, OUTPUT_FOLDER_NAME);  
  //
  // save pdf file
  // save file name pattern : studentId_문서번호
  var pdfName = data.StudentId + '_' + data.문서번호;
  var pdf = createPDF(pdfFolder, document.getId(), pdfName);
  template_copy.setTrashed(true);
  //
  return pdf.getUrl();
}

/**
 * Creates a PDF for the customer given sheet.
 * @param {Object} pdfFolder pdf file saved folder
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created : studentId_roomNumberCode
 * @return {file object} PDF file as a blob
 */
function createPDF(pdfFolder, docsId, pdfName) {
  // const fr = 0, fc = 0, lc = 9, lr = 27;
  // const fr = 0, fc = 0, lc = 0, lr = 29;
  const url = "https://docs.google.com/document/d/" + docsId + "/export" +
    "?format=pdf&" +
    "size=a4&" +          // paper A4 
    "fzr=true&" +         // do not repeat row headers
    "portrait=false&" +   // landscape
    "fitw=true&" +        // fit to page width
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.30&" +
    "bottom_margin=0.00&" +
    "left_margin=0.60&" +
    "right_margin=0.00&" +
    "sheetnames=false&" +
    "pagenum=false&" +
    "attachment=true&"; /**  +
    "gid=" + sheet.getSheetId();
    */
    /** 
     * + "&r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;
     */
  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');
  // Gets the folder in Drive where the PDFs are stored.
  return pdfFolder.createFile(blob);
}

/**
 * data Object Structure
 * 
 * {{문서번호}}
 * {{입주일}}
 * ({{주소}})
 * {{Address}}
 * {{Move-in Date}}
 * {{Name}}
 * {{Student ID Number}}
 * {{생년월일}}
 * {{Birth}}
 * {{발급일자}}
 */
// Helper function to inject data into the template
function populateTemplate(document, data) {

  // clear template
  // 
  // Get the document header and body (which contains the text we'll be replacing).
  var document_header = document.getHeader();
  var document_body = document.getBody();

  // Replace variables in the header
  for (var key in data) {
    var match_text = `{{${key}}}`;
    var value = data[key];

    // Replace our template with the final values
    if(document_header) {
      document_header.replaceText(match_text, value);
    }
    document_body.replaceText(match_text, value);    
  }
}

/**
 * Returns a Google Drive folder in the same location 
 * in Drive where the spreadsheet is located. First, it checks if the folder
 * already exists and returns that folder. If the folder doesn't already
 * exist, the script creates a new one. The folder's name is set by the
 * "OUTPUT_FOLDER_NAME" variable from the Code.gs file.
 *
 * @param {File} templateFile
 * @param {string} folderName - Name of the Drive folder. 
 * @return {object} Google Drive Folder
 */
function getFolderByName_(templateFile, folderName) {
  //
  const parentFolder = templateFile.getParents().next();
  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by ${APP_TITLE} application to store PDF output files`);
}

// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}
