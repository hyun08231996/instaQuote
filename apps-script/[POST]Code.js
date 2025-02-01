/**
 * 아래 doPost 코드는 화면에서 'Save' 또는 'Update' 버튼 클릭 시 작동되며, 입력된 데이터를 'RFQ 리스트' 시트에 넣어준다.
 */

function doPost (e) {
  // lock 기능은 유저들이 QQ에서 동시에 submit(Update/Save) 했을 때, RFQ 데이터가 RFQ 리스트에 1개씩 순차적으로 저장되도록 함 | When users submit (Update/Save) simultaneously in QQ, the lock function ensures that RFQ data is sequentially saved to the RFQ list, one at a time.
  const lock = LockService.getScriptLock(); //모든 사용자가 코드 섹션을 동시에 실행하지 못하도록 잠금 | Gets a lock that prevents any user from concurrently running a section of code.
  lock.tryLock(10000); //잠금을 획득하려고 시도하며 10초 이후 타임아웃됩니다. | Attempts to acquire the lock, timing out after 10 seconds.
 
  try {
    const rfqType = e.parameter['RFQ type'];
    //const rfqType = 'KR';

    const formType = e.parameter['Last submit type'];
    //const formType = 'save';

    //use for Live Link
    const sheetName = rfqType === 'OS' ? formType === 'save-draft' ? 'RFQ List Draft_OS' : 'RFQ List_OS' : formType === 'save-draft' ? 'RFQ List Draft' : 'RFQ List';
    const spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk";

    const doc = SpreadsheetApp.openById(spreadSheetId);
    const sheet = doc.getSheetByName(sheetName);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var totalIdxArr = headers.map((x, idx) => ['Status', 'Last updated', 'Notes', 'Total programming fee', 'Total overlay fee', 'Total other fee', 'Total sales', 'Total GM', 'Total GM (%)', 'Output URL'].includes(x) ? idx + 1 : '').filter(e => e);
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    const lastRange = sheet.getRange(lastRow, 1);
    const lastRFQId = lastRange.isPartOfMerge() ? lastRange.getMergedRanges()[0].getValue() : sheet.getRange(lastRow, 1).getValue() === 'RFQ ID' || sheet.getRange(lastRow, 1).getValue() === 'Draft ID' ? 0 : sheet.getRange(lastRow, 1).getValue();

    //const thisRowVal = sheet.getRange(thisRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    var form = HtmlService.createTemplateFromFile('OpenUrl');

    function saveData(exportURL) {

      var numRows = 1;
      var newRow = [];

      if(rfqType == 'KR') {
        // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
        const addRow = headers.map(function (header) {
          // header가 'RFQ ID'인 경우 마지막 RFQ ID 값에 1을 더함, 그 외의 경우는 e.parameters[header] 사용 | If the header is 'RFQ ID', add 1 to the last RFQ ID value; otherwise, use e.parameters[header]
          return header === 'RFQ ID' ? lastRFQId + 1 : header === 'Status' ? 'Bidding' : e.parameters[header];
        });
        newRow.push(addRow);
        //newRow.push([1,2,3]);
        Logger.log(nextRow);
        Logger.log(numRows);
        

        // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
        //sheet.getRange(nextRow, 1, numRows, headers.length).setValues(newRow);
      }

      if(rfqType == 'OS') {
        var countryCodeList = e.parameter['Country code list[]'].split(',');
        var outputURL = exportURL || '';
        numRows = countryCodeList.length;
        Logger.log(outputURL)

        countryCodeList.forEach(function(code) {
          // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
          const addRow = headers.map(function (header) {
            // header가 'RFQ ID'인 경우 마지막 RFQ ID 값에 1을 더함, 그 외의 경우는 e.parameters[header] 사용 | If the header is 'RFQ ID', add 1 to the last RFQ ID value; otherwise, use e.parameters[header]
            return header === 'RFQ ID' ? lastRFQId + 1 : header === 'GID' ? code : header === 'Status' ? 'Bidding' : header === 'Output URL' ? outputURL : (header === 'Total programming fee' || header === 'Total overlay fee' || header === 'Total other fee' || header === 'Total sales' || header === 'Total GM' || header === 'Total GM (%)') ? e.parameters[header] : e.parameters[header + '-' + code];
            //return header === 'RFQ ID' ? lastRFQId + 1 : e.parameters[header + '-' + code];
          });
          newRow.push(addRow);
          //newRow.push([1,2,3]);
        });
        Logger.log(newRow);
        
      }

      // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
      sheet.getRange(nextRow, 1, numRows, headers.length).setValues(newRow);
      //sheet.getRange(nextRow, 1, numRows, 3).setValues(newRow);

      // merge RFQ ID and Total if multiple rows are saved
      if(rfqType == 'OS' && numRows > 1) {
        sheet.getRange(nextRow, 1, numRows, 1).mergeVertically();
        totalIdxArr.forEach(x => sheet.getRange(nextRow, x, numRows, 1).mergeVertically());
      }
    }

    function saveDraft() {

      var numRows = 1;
      var newRow = [];

      if(rfqType == 'KR') {
        // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
        const addRow = headers.map(function (header) {
          // header가 'RFQ ID'인 경우 마지막 RFQ ID 값에 1을 더함, 그 외의 경우는 e.parameters[header] 사용 | If the header is 'RFQ ID', add 1 to the last RFQ ID value; otherwise, use e.parameters[header]
          return header === 'Draft ID' ? lastRFQId + 1 : e.parameters[header];
        });
        newRow.push(addRow);

        // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
        //sheet.getRange(nextRow, 1, numRows, headers.length).setValues(newRow);
      }

      if(rfqType == 'OS') {
        var countryCodeList = e.parameter['Country code list[]'].split(',');
        numRows = countryCodeList.length;

        countryCodeList.forEach(function(code) {
          // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
          const addRow = headers.map(function (header) {
            // header가 'RFQ ID'인 경우 마지막 RFQ ID 값에 1을 더함, 그 외의 경우는 e.parameters[header] 사용 | If the header is 'RFQ ID', add 1 to the last RFQ ID value; otherwise, use e.parameters[header]
            return header === 'Draft ID' ? lastRFQId + 1 : header === 'GID' ? code : (header === 'Total programming fee' || header === 'Total overlay fee' || header === 'Total other fee' || header === 'Total sales' || header === 'Total GM' || header === 'Total GM (%)') ? e.parameters[header] : e.parameters[header + '-' + code];
            //return header === 'RFQ ID' ? lastRFQId + 1 : e.parameters[header + '-' + code];
          });
          newRow.push(addRow);
          //newRow.push([1,2,3]);
        });
        Logger.log(newRow);
        
      }

      // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
      sheet.getRange(nextRow, 1, numRows, headers.length).setValues(newRow);
      //sheet.getRange(nextRow, 1, numRows, 3).setValues(newRow);

      // merge RFQ ID and Total if multiple rows are saved
      if(rfqType == 'OS' && numRows > 1) {
        sheet.getRange(nextRow, 1, numRows, 1).mergeVertically();
        totalIdxArr.forEach(x => sheet.getRange(nextRow, x, numRows, 1).mergeVertically());
      }
    }

    function updateData() {
      var numRows = 1;
      var updatedRow = [];
      const thisRFQId = e.parameter['RFQ ID'];
      const thisRow = sheet.getRange('A1:A').createTextFinder(thisRFQId).matchEntireCell(true).findNext().getRow();
      const currentRowValues = sheet.getRange(thisRow, 1, 1, headers.length).getValues()[0]; // Fetch the current row's values

      if(rfqType == 'KR') {
        // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
        const updateRow = headers.map(function (header,index) {
          // header가 'RFQ ID'인 경우 마지막 RFQ ID 값에 1을 더함, 그 외의 경우는 e.parameters[header] 사용 | If the header is 'RFQ ID', add 1 to the last RFQ ID value; otherwise, use e.parameters[header]
          return header === 'Last updated' ? new Date() : (header === 'RFQ ID' || header === 'Status' || header === 'Notes') ? currentRowValues[index] : e.parameters[header];
        });
        updatedRow.push(updateRow);

        // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
        //sheet.getRange(nextRow, 1, numRows, headers.length).setValues(newRow);
      }

      if(rfqType == 'OS') {
        var countryCodeList = e.parameter['Country code list[]'].split(',');
        numRows = countryCodeList.length;
        const numMergedRows = sheet.getRange(thisRow, 1).isPartOfMerge() ? sheet.getRange(thisRow, 1).getMergedRanges()[0].getNumRows() : 1;

        if(numRows < numMergedRows) {
          const diff = numMergedRows - numRows;
          for(let i=0; i < diff; i++){
            sheet.deleteRow(thisRow + (numMergedRows - 1) - i);
          }
        } else if (numRows > numMergedRows) {
          const diff = numRows - numMergedRows;
          for(let i=0; i < diff; i++){
            sheet.insertRowAfter(thisRow + (numMergedRows - 1) + i);
          }
        }

        countryCodeList.forEach(function(code) {
          // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
          const updateRow = headers.map(function (header,index) {
            // header가 'RFQ ID'인 경우 마지막 RFQ ID 값에 1을 더함, 그 외의 경우는 e.parameters[header] 사용 | If the header is 'RFQ ID', add 1 to the last RFQ ID value; otherwise, use e.parameters[header]
            return header === 'Last updated' ? new Date() : header === 'GID' ? code : (header === 'Total programming fee' || header === 'Total overlay fee' || header === 'Total other fee' || header === 'Total sales' || header === 'Total GM' || header === 'Total GM (%)') ? e.parameters[header] : (header === 'RFQ ID' || header === 'Status' || header === 'Notes' || header === 'Output URL') ? currentRowValues[index] : e.parameters[header + '-' + code];
            //return header === 'RFQ ID' ? lastRFQId + 1 : e.parameters[header + '-' + code];
          });
          updatedRow.push(updateRow);
          //updatedRow.push([thisRFQId,'Bidding',code]);
        });
        
      }

      // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
      sheet.getRange(thisRow, 1, numRows, headers.length).setValues(updatedRow);
      //sheet.getRange(nextRow, 1, numRows, 3).setValues(newRow);

      // merge RFQ ID and Total if multiple rows are saved
      if(rfqType == 'OS' && numRows > 1) {
        sheet.getRange(thisRow, 1, numRows, 1).mergeVertically();
        totalIdxArr.forEach(x => sheet.getRange(thisRow, x, numRows, 1).mergeVertically());
      }
    }

    function exportData() {

      var countryCodeList = e.parameter['Country code list[]'].split(',');
      var exportFormList = e.parameter['Export form list'].split('||');
      var projectName = e.parameters['Project name (Mail title)-' + countryCodeList[0]];
      var projectType = e.parameters['Project type-' + countryCodeList[0]];
      //var countryCodeList = ['kr','jp'];

      var rows = countryCodeList.length;
      var newRow = [];
      var countriesNameList = [];
      var getValue = (part, o) => Object.entries(o).find(([k, v]) => k.startsWith(part))?.[1];

      countryCodeList.forEach(function(code) {
        // 저장할 행의 각 열 값을 설정 | Set the value of each column in the row to be saved
        const addRow = [e.parameters['Country-' + code],'','',e.parameters['Feasibility-' + code],e.parameters['Proposal CPI-' + code]];
        newRow.push(addRow);
        if (!countriesNameList.includes(e.parameters['Country-' + code])) countriesNameList.push(e.parameters['Country-' + code]);
        //newRow.push([1,2,3])
        //countriesNameList.push('test')
      });

      var countriesName = countriesNameList.length > 5 ? countriesNameList.length + ' countries' : countriesNameList.join(', ');

      var specificFolder = DriveApp.getFolderById('1O4u0rnwU5ukORLJTLyD58sv6aR0xgXr9');
      var date = new Date();
      var formattedDate = getFormatDate(date);
      var file = DriveApp.getFileById('1i9Q4xmM7r_iTxcPe1ZH38rzqz6hmLfsrWybYZbgr8yY');
      var fileCopy = file.makeCopy(`${ formattedDate }_${ exportFormList[0] }_${ countriesName }_${ projectName }`, specificFolder);
      var fileCopyId = fileCopy.getId();
      var fileCopyUrl = fileCopy.getUrl();

      const doc = SpreadsheetApp.openById(fileCopyId);
      const sheet = doc.getSheetByName('Quotation');

      var rowIndex = 22; // Insert rows after row 5
      var numRows = rows - 2;  // Number of rows to insert
      var sourceRange = sheet.getRange("B"+ rowIndex +":G"+ rowIndex);
      

      sheet.getRange('B13').setValue(exportFormList[0]);
      sheet.getRange('B12').setValue(exportFormList[1]);
      sheet.getRange('B4').setValue(exportFormList[2]);
      sheet.getRange('D15').setValue(exportFormList[3] + 'ss/' + exportFormList[4] + ' min');
      sheet.getRange('F19').setValue(exportFormList[5]);
      sheet.getRange('E20').setValue(exportFormList[6]);
      sheet.getRange('G20').setValue(exportFormList[7]);
      sheet.getRange('E21').setValue(exportFormList[8]);
      sheet.getRange('G21').setValue(exportFormList[9]);
      sheet.getRange('D12').setValue(exportFormList[10]);
      sheet.getRange('F12').setValue(exportFormList[11]);
      
      if(numRows > 0){

        sheet.insertRowsAfter(rowIndex,numRows);
        
        // Loop to copy formatting to each newly inserted row
        for (var i = 1; i <= numRows; i++) {
          var destinationRow = rowIndex + i;
          // Get the range of the newly inserted row
          var destinationRange = sheet.getRange("B"+ destinationRow +":G"+ destinationRow);
          Logger.log(destinationRow);
          
          // Copy the format from the row above to the newly inserted row
          sourceRange.copyTo(destinationRange);
          
        }

      }
      // 저장된 행의 값을 시트에 설정 | Set the values of the saved row in the sheet
      sheet.getRange(rowIndex, 2, rows, 5).setValues(newRow);

      return fileCopyUrl;
    }

    // RFQ List에서 'RFQ Status' 변경 시.. | When 'Update' button is clicked in QQ..
    if (formType == 'update-status') {
      const thisRFQId = e.parameter['RFQ ID'];
      const thisRow = sheet.getRange('A1:A').createTextFinder(thisRFQId).matchEntireCell(true).findNext().getRow();
      const changedStatus = e.parameter['RFQ Status'];
      const statusColumnIndex = headers.indexOf('Status') + 1; // Find the 'Status' column
      const lastUpdatedColumnIndex = headers.indexOf('Last updated') + 1; // Find the 'Last updated' column

      sheet.getRange(thisRow, statusColumnIndex, 1, 1).setValue(changedStatus);
      sheet.getRange(thisRow, lastUpdatedColumnIndex, 1, 1).setValue(new Date());
    }

    // if (formType == 'update-crmno') {
    //   const thisRFQId = e.parameter['RFQ ID'];
    //   const thisRow = sheet.getRange('A1:A').createTextFinder(thisRFQId).matchEntireCell(true).findNext().getRow();
    //   const changedCRMNo = e.parameter['CRM No.'];
    //   const crmNoColumnIndex = headers.indexOf('CRM No.') + 1; // Find the 'CRM No.' column
    //   const lastUpdatedColumnIndex = headers.indexOf('Last updated') + 1; // Find the 'Last updated' column

    //   sheet.getRange(thisRow, crmNoColumnIndex, 1, 1).setValue(changedCRMNo);
    //   sheet.getRange(thisRow, lastUpdatedColumnIndex, 1, 1).setValue(new Date());
    // }

    if (formType == 'update-outputurl') {
      const thisRFQId = e.parameter['RFQ ID'];
      const thisRow = sheet.getRange('A1:A').createTextFinder(thisRFQId).matchEntireCell(true).findNext().getRow();
      const changedCRMNo = e.parameter['Output URL'];
      const crmNoColumnIndex = headers.indexOf('Output URL') + 1; // Find the 'Output URL' column
      const lastUpdatedColumnIndex = headers.indexOf('Last updated') + 1; // Find the 'Last updated' column

      sheet.getRange(thisRow, crmNoColumnIndex, 1, 1).setValue(changedCRMNo);
      sheet.getRange(thisRow, lastUpdatedColumnIndex, 1, 1).setValue(new Date());
    }

    if (formType == 'update-comments') {
      const thisRFQId = e.parameter['RFQ ID'];
      const thisRow = sheet.getRange('A1:A').createTextFinder(thisRFQId).matchEntireCell(true).findNext().getRow();
      const changedComments = e.parameter['Notes'];
      const commentsColumnIndex = headers.indexOf('Notes') + 1; // Find the 'Comments' column
      const lastUpdatedColumnIndex = headers.indexOf('Last updated') + 1; // Find the 'Last updated' column

      sheet.getRange(thisRow, commentsColumnIndex, 1, 1).setValue(changedComments);
      sheet.getRange(thisRow, lastUpdatedColumnIndex, 1, 1).setValue(new Date());
    }

    // QQ에서 'Update' 버튼 클릭 시.. | When 'Update' button is clicked in QQ..
    if (formType == 'update') {
      updateData();
    }

    // QQ에서 'Save' 버튼 클릭 시.. | When 'Save' button is clicked in QQ..
    if (formType == 'save') {
      saveData();
    }

    if (formType == 'save-draft') {
      saveDraft();
    }

    if (formType == 'export') {
      var exportURL = exportData();
      form.url = exportURL;
      
      return form.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Ensure the page is allowed to open new tabs
    }

    if (formType == 'save-export') {
      var exportURL = exportData();
      saveData(exportURL);
      form.url = exportURL;
      
      return form.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Ensure the page is allowed to open new tabs
    }
    
    // return ContentService
    //   .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
    //   .setMimeType(ContentService.MimeType.JSON);
    //return HtmlService.createHtmlOutputFromFile('[HTML]Success.html').setTitle('QuickQ (QQ) - Success');
  }
  
  // 오류 발생 시 처리 | Handle errors if they occur
  catch (err) {
    return err;
    // return ContentService
    //   .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
    //   .setMimeType(ContentService.MimeType.JSON);
  }
 
  finally {
    lock.releaseLock(); //잠금을 해제하여 잠금에서 대기 중인 다른 프로세스가 계속되도록 허용합니다. | Releases the lock, allowing other processes waiting on the lock to continue.
  }
}

/**
 *  YYMMDD 포맷으로 반환
 */
function getFormatDate(date){
    var year = date.getFullYear().toString().slice(-2);    //yy
    var month = (1 + date.getMonth());          //M
    month = month >= 10 ? month : '0' + month;  //month 두자리로 저장
    var day = date.getDate();                   //d
    day = day >= 10 ? day : '0' + day;          //day 두자리로 저장
    return  year + '' + month + '' + day;       
}



