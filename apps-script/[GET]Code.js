/**
 * 아래 doGet 코드는 QQ에 접속하면 바로 실행되며, 'Rate Card', 'Client 리스트', 'Country 리스트' 등 다양한 시트에서 데이터를 불러온다.
 */
eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.30.1/moment.min.js').getContentText());

var numRetries = 6;

var weekStart = moment().startOf('week').format('M/D/YYYY');
var weekEnd = moment().endOf('week').format('M/D/YYYY');
var monthStart = moment().startOf('month').format('M/D/YYYY');
var monthEnd = moment().endOf('month').format('M/D/YYYY');
var quarterStart = moment().startOf('quarter').format('M/D/YYYY');
var quarterEnd = moment().endOf('quarter').format('M/D/YYYY');
var yearStart = moment().quarter() == 1 ? moment().subtract(1,'year').month('April').startOf('month').format('M/D/YYYY') : moment().month('April').startOf('month').format('M/D/YYYY');
var yearEnd = moment().quarter() == 1 ? moment().month('March').endOf('month').format('M/D/YYYY') : moment().add(1,'year').month('March').endOf('month').format('M/D/YYYY');

function doGet(e) {

  // Set Active page
  let page = e.parameter.mode || "Index";

  // 초기 Oauth 매개변수 가져오기, 없으면 false로 설정 | Get initial OAuth parameter, set to false if not present
  const initOauth = (JSON.stringify(e.parameters.oauth)) || false;
  
  // Web 앱의 URL 가져오기 | Get the URL of the web app
  const webAppUrl = ScriptApp.getService().getUrl();

  // 초기 Oauth가 없으면 실행 모드에 따라 적절한 HTML 페이지 반환 | If there is no initial OAuth, return the appropriate HTML page based on the execution mode
  if (!initOauth) {
    var form = HtmlService.createTemplateFromFile(page);

    if(page == 'Index') {
      //var userName = getActiveUserEmail().split('.')[0];
      //var modName = userName[0].toUpperCase() + userName.slice(1);

      form.USER_NAME = getActiveUserName();
    }

    form.SCRIPT_URL = getScriptURL();

    var formEval = form.evaluate();
    var formOutput = HtmlService.createHtmlOutput(formEval);

    //Replace {{NAV_SIDEBAR}} with the getNavSidebar content
    formOutput.setContent(formOutput.getContent().replace("{{NAV_SIDEBAR}}",getNavSidebar(page)).replace("{{FOOTER}}",getFooter()));

    //form.assets = getAssets('견적 폼 이미지');
    // Web 앱 URL이 "/exec"로 끝나면 실행 페이지 반환 | If the web app URL ends with "/exec," return the execution page
    if (webAppUrl.includes("/exec")) {
      return formOutput.setTitle('instaQuote').setFaviconUrl('https://i.postimg.cc/2qPmbrwS/insta-Quote-logo-final.png');
    } 
    // Web 앱 URL이 "/dev"로 끝나면 탭 이름에 'Test' 표시 | If the web app URL ends with "/dev," show 'Test' in the tab name
    else if (webAppUrl.includes("/dev")) {
      return formOutput.setTitle('instaQuote - Test').setFaviconUrl('https://i.postimg.cc/2qPmbrwS/insta-Quote-logo-final.png');
    }
  } 
  // 초기 Oauth가 있으면 OAuth 인증 페이지 반환 | If there is initial OAuth, return the OAuth authentication page
  else {
    return buildOAuthAuthPage(e);
  }
}

function getFooter() {
  return `<footer class="footer">
    <div>Copyright © 2025 Peter Nam.</div>
    <a href="https://script.google.com/macros/s/AKfycbzR4C6JvphST5AY-i8Gzpv1ECapoC0pwl4Xyv8qx4i2Mv8sIO8chVT4fKumq_VCIVr4/exec?mode=privacy-policy" target="_blank">Privacy Policy & Terms of Service</a>
  </footer>`;
}

//Create Navigation Sidebar
function getNavSidebar(activePage) {
  var scriptURLHome = getScriptURL();
  var scriptURLCreate = getScriptURL("mode=Create-KR");
  var scriptURLDash = getScriptURL("mode=Dash-KR");
  var scriptURLStatus = getScriptURL("mode=Status-KR");
  var scriptURLList = getScriptURL("mode=List-KR");
  var scriptURLVendor = getScriptURL("mode=Vendor-OS");
  var scriptURLDraft = getScriptURL("mode=Draft-KR");
  //var scriptURLPage2 = getScriptURL("mode=Page2");

  var navbar = 
    `<div id="navSidebar" class="sidebar" onmouseenter="toggleSidebar()" onmouseleave="toggleSidebar()">
      <div class="sidebar-no-overflow-x">
        <!--<center>-->
          <div class="navbar-brand">
            <a id="navbarLogo" target="_top" href="${scriptURLHome}" style="text-decoration:none;">
              <img class="navbar-brand-logo" src="https://i.postimg.cc/2qPmbrwS/insta-Quote-logo-final.png" />
              <div class="navbar-brand-text" style="font-weight:bolder;"><span style="color:white;">insta</span><span style="color:#9270ff;">Quote</span></div>
            </a>
          </div>
        
          <div class="btn-group rfq-group" role="group" aria-label="Basic radio toggle button group">
            <input type="radio" class="btn-check rfq-type-check" name="btnradio" id="rfqTypeKR" autocomplete="off" value="KR" checked>
            <label class="btn btn-outline-secondary rfq-type-kr" for="rfqTypeKR">KR</label>

            <input type="radio" class="btn-check rfq-type-check" name="btnradio" id="rfqTypeOS" autocomplete="off" value="OS">
            <label class="btn btn-outline-secondary rfq-type-os" for="rfqTypeOS">OS</label>

            <div class="btn-slider"></div>
          </div>
          
          <div class="nav-menu-link-container">
            <a class="btn btn-primary nav-menu-link create-rfq-btn" target="_top" href="${scriptURLCreate}" role="button"><i class="bi bi-plus-square"></i><span class="bi-icon-text">&nbsp;&nbsp;Create RFQ</span></a>
          </div>

          <div class="nav-menu-container">
            <div class="nav-menu-link-container">
              <a class="btn nav-menu-link ${activePage.includes('Draft') ? 'active' : ''}" target="_top" href="${scriptURLDraft}" role="button"><i class="bi bi-table"></i></i><span class="bi-icon-text">&nbsp;&nbsp;Drafts</span></a>
            </div>
            <div class="nav-menu-link-container">
              <a class="btn nav-menu-link ${activePage.includes('List') ? 'active' : ''}" target="_top" href="${scriptURLList}" role="button"><i class="bi bi-table"></i><span class="bi-icon-text">&nbsp;&nbsp;RFQ List</span></a>
            </div>
            <div class="nav-menu-link-container">
              <a class="btn nav-menu-link ${activePage.includes('Dash') ? 'active' : ''}" target="_top" href="${scriptURLDash}" role="button"><i class="bi bi-speedometer2"></i><span class="bi-icon-text">&nbsp;&nbsp;RFQ Dashboard</span></a>
            </div>
            <div class="nav-menu-link-container">
              <a class="btn nav-menu-link ${activePage.includes('Status') ? 'active' : ''}" target="_top" href="${scriptURLStatus}" role="button"><i class="bi bi-bar-chart-line"></i><span class="bi-icon-text">&nbsp;&nbsp;Order Status</span></a>
            </div>
            <div class="nav-menu-link-container">
              <a class="btn nav-menu-link ${activePage.includes('Vendor') ? 'active' : ''}" target="_top" href="${scriptURLVendor}" role="button"><i class="bi bi-table"></i></i><span class="bi-icon-text">&nbsp;&nbsp;Vendor List</span></a>
            </div>
            
          </div>
        <!--</center>-->

        
          
        
      </div>

      <div class="link-menu-dropdown-container">
        <div class="dropdown link-menu-dropdown dropup">
          <button class="btn link-menu" data-bs-toggle="dropdown" data-bs-hover="dropdown" data-bs-auto-close="true" aria-expanded="false">
              <i class="bi bi-translate" style="color: white;"></i>
          </button>
          <div class="dropdown-menu dropdown-menu-top nav-dropdown-menu nav-dropdown-menu-end shadow-sm">
            <a class="dropdown-item nav-dropdown-item" href="#" id="switchToKorean">한국어</a>
            <a class="dropdown-item nav-dropdown-item" href="#" id="switchToEnglish">English</a>
          </div>
        </div>
        <div class="dropdown link-menu-dropdown dropup">
          <button class="btn link-menu" data-bs-toggle="dropdown" data-bs-hover="dropdown" aria-expanded="false">
            <i class="bi bi-box-arrow-up-right" style="color:white;"></i>
          </button>
          <div class="dropdown-menu dropdown-menu-top nav-dropdown-menu shadow-sm" style="width:500px;height:300px;padding:0;">
            <div class="row" style="width:100%;height:100%;margin:0;">
              <div class="col-5" style="background:#f5f3f8;border-radius:0.25rem 0 0 0.25rem;padding:0.5rem;">
                <h4 class="dropdown-header">RFQ Related</h4>
                <a class="dropdown-item nav-dropdown-item" href="https://docs.google.com/spreadsheets/d/1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk/edit?gid=0#gid=0" target="_blank"><div class="shadow-sm"><img src="https://ssl.gstatic.com/docs/spreadsheets/spreadsheets_2023q4.ico" height="15px" width="15px"></div>RFQ List (KR)</a>
                <a class="dropdown-item nav-dropdown-item" href="https://docs.google.com/spreadsheets/d/1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk/edit?gid=36461421#gid=36461421" target="_blank"><div class="shadow-sm"><img src="https://ssl.gstatic.com/docs/spreadsheets/spreadsheets_2023q4.ico" height="15px" width="15px"></div>RFQ List (OS)</a>
                <a class="dropdown-item nav-dropdown-item" href="https://docs.google.com/spreadsheets/d/1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk/edit?gid=1004620658#gid=1004620658" target="_blank"><div class="shadow-sm"><img src="https://ssl.gstatic.com/docs/spreadsheets/spreadsheets_2023q4.ico" height="15px" width="15px"></div>Vendor List</a>
                <a class="dropdown-item nav-dropdown-item" href="https://docs.google.com/spreadsheets/d/1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk/edit?gid=263916741#gid=263916741" target="_blank"><div class="shadow-sm"><img src="https://ssl.gstatic.com/docs/spreadsheets/spreadsheets_2023q4.ico" height="15px" width="15px"></div>Country Code</a>
              </div>
              <div class="col-7" style="padding:0.5rem;">
                <h4 class="dropdown-header">Useful Links</h4>
                <div id="insertLinkList" style="height:225px;overflow-y:auto;"></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>`;
  return navbar;
}

function getSheetNameById(spreadSheetId,gid) {
  var url = SpreadsheetApp.openById(spreadSheetId).getUrl();
  var ss = SpreadsheetApp.openByUrl(url);
  var sheetId = gid;
  var sheet = ss.getSheets().find(sheet => sheet.getSheetId() == sheetId);

  return sheet.getSheetName();
}
 
//GET DATA AND RETURN AS AN ARRAY
function getRate(){
  
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  //Logger.log(getSheetNameById(spreadSheetId,'869037984'))
  var dataRange0     = getSheetNameById(spreadSheetId,'358935832') + "!C3:P14"; //Group A
  var dataRange1     = getSheetNameById(spreadSheetId,'358935832') + "!C17:P28"; //Group B
  var dataRange2     = getSheetNameById(spreadSheetId,'358935832') + "!C31:P42"; //Group C
  var dataRange3     = getSheetNameById(spreadSheetId,'358935832') + "!C45:P56"; //Group D
  var dataRange4     = getSheetNameById(spreadSheetId,'358935832') + "!C59:P70"; //Group E
 
  // 해당하는 구글 시트에서 데이터를 가져옴 | Get data from the corresponding Google Sheets
  var range0   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange0);
  var range1   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange1);
  var range2   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange2);
  var range3   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange3);
  var range4   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange4);
  // var range5   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange5);
  // var range6   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange6);
  // var range7   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange7);
  // var range8   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange8);
  // var range9   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange9);
  // var range10   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange10);

  var ranges = [range0,range1,range2,range3,range4];
  
  var values0  = range0.values;
  var values1  = range1.values;
  var values2  = range2.values;
  var values3  = range3.values;
  var values4  = range4.values;
  // var values5  = range5.values;
  // var values6  = range6.values;
  // var values7  = range7.values;

  // var data = range8.values;
  // var values8  = [];
  // var values9  = [];
  // var values10 = [];
  // var values11 = [];
  // var values12 = [];
  // var values13 = [];
  // var values14  = [];

  // var values15 = range9.values;
  // var values16 = range10.values;

  // for(i in data){
  //   // 데이터를 해당하는 범위에 따라 분리하여 값을 할당 | Split the data and assign values based on the corresponding range
  //   if(i>=0 && i<=5){ for(var j=0;j<=13;j++){ values8.push(data[i][j]); } }
  //   if(i>=7 && i<=12){ for(var j=0;j<=13;j++){ values9.push(data[i][j]); } }
  //   if(i>=14 && i<=19){ for(var j=0;j<=13;j++){ values10.push(data[i][j]); } }
  //   if(i>=21 && i<=26){ for(var j=0;j<=13;j++){ values11.push(data[i][j]); } }
  //   if(i>=28 && i<=33){ for(var j=0;j<=13;j++){ values12.push(data[i][j]); } }
  //   if(i>=35 && i<=40){ for(var j=0;j<=13;j++){ values13.push(data[i][j]); } }
  //   if(i>=42 && i<=47){ for(var j=0;j<=13;j++){ values14.push(data[i][j]); } }
  // }

  // // 데이터를 14열로 나누어서 할당 | Split the data into chunks of 14 columns and assign values
  // values8 = spliceIntoChunks(values8,14);
  // values9 = spliceIntoChunks(values9,14);
  // values10 = spliceIntoChunks(values10,14);
  // values11 = spliceIntoChunks(values11,14);
  // values12 = spliceIntoChunks(values12,14);
  // values13 = spliceIntoChunks(values13,14);
  // values14 = spliceIntoChunks(values14,14);

  Logger.log(ranges);

  // 스프레드시트를 가져오기 위한 함수 | Function to get the spreadsheet
  var getSpreadsheet = function() {return ranges};
  // 리트라이를 사용하여 스프레드시트 가져오기 시도 | Attempt to get the spreadsheet using retry
  retry_(getSpreadsheet, numRetries);
  //Logger.log(values14);
  // 모든 값 반환 | Return all values
  return [values0,values1,values2,values3,values4];
}

//GET DATA FROM 'Client 리스트' SHEET AND RETURN AS AN ARRAY
function getClient(){
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var dataRange     = "Client List!A:XX"; //CHANGE
 
  var range   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange);
  var data_full  = range.values;
  var header = data_full[0];
  var data = data_full.slice(2);

  var valuesH = [];
  var values0  = [];
  var values1  = [];
  var values2  = [];
  var values3  = [];
  var values4  = [];
  // var values5  = [];
  // var values6  = [];
  // var values7  = [];
  // var values8  = [];

  header.forEach(val => {
    if(val !== '') valuesH.push(val);
  });

  for(i in data){
    if(data[i][0] != null && data[i][1] != null && data[i][0] != '' && data[i][1] != ''){ values0.push([data[i][0], data[i][1]]); }
    if(data[i][2] != null && data[i][3] != null && data[i][2] != '' && data[i][3] != ''){ values1.push([data[i][2], data[i][3]]); }
    if(data[i][4] != null && data[i][5] != null && data[i][4] != '' && data[i][5] != ''){ values2.push([data[i][4], data[i][5]]); }
    if(data[i][6] != null && data[i][7] != null && data[i][6] != '' && data[i][7] != ''){ values3.push([data[i][6], data[i][7]]); }
    if(data[i][8] != null && data[i][9] != null && data[i][8] != '' && data[i][9] != ''){ values4.push([data[i][8], data[i][9]]); }
    // if(data[i][10] != null && data[i][11] != null && data[i][10] != '' && data[i][11] != ''){ values5.push([data[i][10], data[i][11]]); }
    // if(data[i][12] != null && data[i][13] != null && data[i][12] != '' && data[i][13] != ''){ values6.push([data[i][12], data[i][13]]); }
    // if(data[i][14] != null && data[i][15] != null && data[i][14] != '' && data[i][15] != ''){ values7.push([data[i][14], data[i][15]]); }
    // if(data[i][16] != null && data[i][17] != null && data[i][16] != '' && data[i][17] != ''){ values8.push([data[i][16], data[i][17]]); }
  }

  var getSpreadsheet = function() {return range};
  retry_(getSpreadsheet, numRetries);
  Logger.log(valuesH);
  return [valuesH,values0,values1,values2,values3,values4];
}

//GET DATA FROM '참고 링크' SHEET AND RETURN AS AN ARRAY
function getLink(){
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var dataRange     = "Useful Links!A2:B"; //CHANGE
 
  var range   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange);
  
  var values  = range.values;

  
  var getSpreadsheet = function() {return range};
  retry_(getSpreadsheet, numRetries);
 
  return values;
}

//GET DATA FROM 'Country 리스트' SHEET AND RETURN AS AN ARRAY
function getCountry(){
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var dataRange     = "Country List!A2:E"; //CHANGE
 
  var range   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange);
  
  var values  = range.values;

  
  var getSpreadsheet = function() {return range};
  retry_(getSpreadsheet, numRetries);
 
  return values;
}

//GET DATA FROM 'Data' SHEET AND RETURN AS AN ARRAY
function getCompPt(){
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var dataRange     = "Comp Pt!A2:B"; //CHANGE
 
  var range   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange);
  
  var values  = range.values;

  
  var getSpreadsheet = function() {return range};
  retry_(getSpreadsheet, numRetries);
 
  return values;
}

//GET DATA FROM '운영비' SHEET AND RETURN AS AN ARRAY
function getOtherFee(){
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var dataRange     = "Other Fee!A3:O"; //CHANGE
 
  var range   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange);

  var data = range.values;
  
  var values0  = [];
  var values1  = [];
  var values2  = [];
  var values3  = [];
  var values4  = [];

  for(i in data){
    if(data[i][0] != null && data[i][1] != null && data[i][2] != null && data[i][0] != '' && data[i][1] != '' && data[i][2] != ''){ values0.push([data[i][0], data[i][1], data[i][2]]); }
    if(data[i][3] != null && data[i][4] != null && data[i][5] != null && data[i][3] != '' && data[i][4] != '' && data[i][5] != ''){ values1.push([data[i][3], data[i][4], data[i][5]]); }
    if(data[i][6] != null && data[i][7] != null && data[i][8] != null && data[i][6] != '' && data[i][7] != '' && data[i][8] != ''){ values2.push([data[i][6], data[i][7], data[i][8]]); }
    if(data[i][9] != null && data[i][10] != null && data[i][11] != null && data[i][9] != '' && data[i][10] != '' && data[i][11] != ''){ values3.push([data[i][9], data[i][10], data[i][11]]); }
    if(data[i][12] != null && data[i][13] != null && data[i][14] != null && data[i][12] != '' && data[i][13] != '' && data[i][14] != ''){ values4.push([data[i][12], data[i][13], data[i][14]]); }
  }

  var getSpreadsheet = function() {return range};
  retry_(getSpreadsheet, numRetries);
 
  return [values0,values1,values2,values3,values4];
}

//use for Live Link
//GET DATA FROM 'RFQ 리스트' SHEET AND RETURN AS AN ARRAY
function getRFQ(rfqid){
  //use for Live Link
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var sheetName = "RFQ List";

  var doc = SpreadsheetApp.openById(spreadSheetId);
  var sheet = doc.getSheetByName(sheetName);
  
  //var rangeValues = sheet.getDataRange();
  //var values = rangeValues.getValues();
  //var values = sheet.getDataRange().getDisplayValues();
  var values;

  if(rfqid) {
    var thisRFQRow = sheet.getRange('A1:A').createTextFinder(rfqid).matchEntireCell(true).findNext().getRow();
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var thisRFQValues = sheet.getRange(thisRFQRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    values = [header,thisRFQValues];
    //values.filter((arr,idx) => arr[0].toString() == rfqid.toString() || idx == 0);
  } else {
    values = sheet.getDataRange().getDisplayValues();
  }

  Logger.log(values)
  
  var getSpreadsheet = function() {return values};
  retry_(getSpreadsheet, numRetries);
 
  return values;
}

function getDraft(rfqid){
  //use for Live Link
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var sheetName = "RFQ List Draft";


  var doc = SpreadsheetApp.openById(spreadSheetId);
  var sheet = doc.getSheetByName(sheetName);

  var values;

  if(rfqid) {
    var thisRFQRow = sheet.getRange('A1:A').createTextFinder(rfqid).matchEntireCell(true).findNext().getRow();
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var thisRFQValues = sheet.getRange(thisRFQRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    values = [header,thisRFQValues];
    //values.filter((arr,idx) => arr[0].toString() == rfqid.toString() || idx == 0);
  } else {
    values = sheet.getDataRange().getDisplayValues();
  }

  Logger.log(values)
  
  var getSpreadsheet = function() {return values};
  retry_(getSpreadsheet, numRetries);
 
  return values;
}

function getRFQOS(rfqid){
  //use for Live Link
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var sheetName = "RFQ List_OS";


  var doc = SpreadsheetApp.openById(spreadSheetId);
  var sheet = doc.getSheetByName(sheetName);

  // var rangeValues = sheet.getDataRange();
  // //var values = rangeValues.getValues();
  // var values = rangeValues.getDisplayValues();
  //var values = sheet.getDataRange().getDisplayValues();
  var values;

  if(rfqid) {
    var thisRFQRow = sheet.getRange('A1:A').createTextFinder(rfqid).matchEntireCell(true).findNext().getRow();
    const numMergedRows = sheet.getRange(thisRFQRow, 1).isPartOfMerge() ? sheet.getRange(thisRFQRow, 1).getMergedRanges()[0].getNumRows() : 1;
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var thisRFQValues = sheet.getRange(thisRFQRow, 1, numMergedRows, sheet.getLastColumn()).getDisplayValues();

    values = [header,...thisRFQValues];
  } else {
    values = sheet.getDataRange().getDisplayValues();

    // Get merged ranges in the sheet
    //var mergedRanges = sheet.getRange(rangeValues.getA1Notation()).getMergedRanges();
    var mergedRanges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getMergedRanges();

    var mergedInfo = [];

    // Iterate through the merged ranges
    mergedRanges.forEach(function(range) {
      
      
      // // Get the range's row and column information
      // var startRow = range.getRow();
      // var startCol = range.getColumn();
      // var numRows = range.getNumRows();
      // var numCols = range.getNumColumns();

      // // Get the value of the merged range
      // // var mergedValue = range.getCell(1, 1).getValue(); // The value of the top-left cell of the merge
      // var mergedValue = values[startRow][startCol]; // The value of the top-left cell of the merge

      // mergedInfo.push({
      //   row: startRow - 1, 
      //   col: startCol - 1, 
      //   rowspan: numRows, 
      //   colspan: numCols,
      //   mergedValue: mergedValue
      // });

      // Use pre-fetched values instead of calling getCell().getValue()
      var startRow = range.getRow() - 1; // Convert to zero-based index
      var startCol = range.getColumn() - 1; // Convert to zero-based index
      var mergedValue = values[startRow][startCol]; // Use pre-fetched values

      // Collect merged range info
      mergedInfo.push({
        row: startRow,
        col: startCol,
        rowspan: range.getNumRows(),
        colspan: range.getNumColumns(),
        mergedValue: mergedValue
      });

      if(startCol === 0) {
        // Apply the merged value to all cells in the merged range
        for (var row = startRow; row < startRow + range.getNumRows(); row++) {
          for (var col = startCol; col < startCol + range.getNumColumns(); col++) {
            values[row][col] = mergedValue.toString();
          }
        }
      }
    });


    // Reverse all rows except the first (header row)
    var header = values[0]; // Keep the first row (header) intact
    var bodyRows = values.slice(1); // Get all rows after the header
    //var reversedBodyRows = bodyRows.reverse(); // Reverse the body rows
    var reversedBodyRows = [];

    //Logger.log(values)


    // Adjust merged cell info only for body rows
    var rowCount = bodyRows.length; // Row count without the header
    mergedInfo.forEach(function(merge) {
      if (merge.row > 0) { // Only adjust rows after the header (row 0)
        merge.row = rowCount - merge.row  - (merge.rowspan - 2); // Adjust for reversed rows
      }
    });
    //Logger.log(bodyRows)

    var groupedArr = bodyRows.reduce((r, v) => {  
      const id = v[0];

      r[id] = r[id] || [];
      r[id].push(v);

      return r;
    }, {});

    var reversedArr = Object.values(groupedArr).reverse();

    reversedArr.forEach(function(a1) {
      a1.forEach(function(a2) {
        reversedBodyRows.push(a2);
      });
    });

    // Combine the header with the reversed body rows
    var finalValues = [header].concat(reversedBodyRows);

    // Index of the "Last updated" column (adjust index as needed)
    var lastUpdatedColumnIndex = header.indexOf("Last updated");

    // Convert "Last updated" values to ISO strings
    if (lastUpdatedColumnIndex > -1) {
        finalValues.forEach((row, index) => {
            if (index > 0 && row[lastUpdatedColumnIndex]) { // Skip the header row
                var cellValue = row[lastUpdatedColumnIndex];
                try {
                    // Try to parse and convert to ISO string
                    var date = new Date(cellValue);
                    if (!isNaN(date)) {
                        row[lastUpdatedColumnIndex] = moment(date).format('M/D/YYYY');
                    }
                } catch (e) {
                    Logger.log("Error parsing date: " + cellValue);
                }
            }
        });
    }
  }

  

  //Logger.log(finalValues)
  var getSpreadsheet = function() {return values};
  retry_(getSpreadsheet, numRetries);
 
  //return values;
  //return { dataArray: values, mergedInfo: mergedInfo };
  Logger.log(values)
  //return { dataArray: finalValues, mergedInfo: mergedInfo };
  //return finalValues;
  return rfqid ? JSON.parse(JSON.stringify({ dataArray: values })) : JSON.parse(JSON.stringify({ dataArray: finalValues, mergedInfo: mergedInfo }));
}

function getDraftOS(rfqid){
  //use for Live Link
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var sheetName = "RFQ List Draft_OS";


  var doc = SpreadsheetApp.openById(spreadSheetId);
  var sheet = doc.getSheetByName(sheetName);

  // var rangeValues = sheet.getDataRange();
  // //var values = rangeValues.getValues();
  // var values = rangeValues.getDisplayValues();
  var values;

  if(rfqid) {
    var thisRFQRow = sheet.getRange('A1:A').createTextFinder(rfqid).matchEntireCell(true).findNext().getRow();
    const numMergedRows = sheet.getRange(thisRFQRow, 1).isPartOfMerge() ? sheet.getRange(thisRFQRow, 1).getMergedRanges()[0].getNumRows() : 1;
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    var thisRFQValues = sheet.getRange(thisRFQRow, 1, numMergedRows, sheet.getLastColumn()).getDisplayValues();

    values = [header,...thisRFQValues];
  } else {
    values = rangeValues.getDisplayValues();

    // Get merged ranges in the sheet
    //var mergedRanges = sheet.getRange(rangeValues.getA1Notation()).getMergedRanges();
    var mergedRanges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getMergedRanges();

    var mergedInfo = [];

    // Iterate through the merged ranges
    mergedRanges.forEach(function(range) {
      // Get the value of the merged range
      var mergedValue = range.getCell(1, 1).getValue(); // The value of the top-left cell of the merge
      
      // Get the range's row and column information
      var startRow = range.getRow();
      var startCol = range.getColumn();
      var numRows = range.getNumRows();
      var numCols = range.getNumColumns();

      mergedInfo.push({
        row: startRow - 1, 
        col: startCol - 1, 
        rowspan: numRows, 
        colspan: numCols,
        mergedValue: mergedValue
      });
      if(startCol - 1 === 0) {
      // Apply the merged value to all cells in the merged range
      for (var row = startRow; row < startRow + numRows; row++) {
        for (var col = startCol; col < startCol + numCols; col++) {
          values[row - 1][col - 1] = mergedValue.toString();
        }
      }
      }
    });


    // Reverse all rows except the first (header row)
    var header = values[0]; // Keep the first row (header) intact
    var bodyRows = values.slice(1); // Get all rows after the header
    //var reversedBodyRows = bodyRows.reverse(); // Reverse the body rows
    var reversedBodyRows = [];

    Logger.log(values)


    // Adjust merged cell info only for body rows
    var rowCount = bodyRows.length; // Row count without the header
    mergedInfo.forEach(function(merge) {
      if (merge.row > 0) { // Only adjust rows after the header (row 0)
        merge.row = rowCount - merge.row  - (merge.rowspan - 2); // Adjust for reversed rows
      }
    });
    //Logger.log(bodyRows)

    var groupedArr = bodyRows.reduce((r, v) => {  
      const id = v[0];

      r[id] = r[id] || [];
      r[id].push(v);

      return r;
    }, {});

    var reversedArr = Object.values(groupedArr).reverse();

    reversedArr.forEach(function(a1) {
      a1.forEach(function(a2) {
        reversedBodyRows.push(a2);
      });
    });

    // Combine the header with the reversed body rows
    var finalValues = [header].concat(reversedBodyRows);
  }

  

  //Logger.log(finalValues)
  var getSpreadsheet = function() {return rangeValues};
  retry_(getSpreadsheet, numRetries);
 
  //return values;
  //return { dataArray: values, mergedInfo: mergedInfo };
  //return { dataArray: finalValues, mergedInfo: mergedInfo };
  return rfqid ? JSON.parse(JSON.stringify({ dataArray: values })) : JSON.parse(JSON.stringify({ dataArray: finalValues, mergedInfo: mergedInfo }));
}



function getFilteredRFQ() {
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  var sheetName1 = "RFQ List";
  var sheetName2 = "RFQ List_OS";

  var doc = SpreadsheetApp.openById(spreadSheetId);
  var sheet1 = doc.getSheetByName(sheetName1);
  var sheet2 = doc.getSheetByName(sheetName2);

  var rangeValues1 = sheet1.getDataRange();
  var rangeValues2 = sheet2.getDataRange();
  var values1 = rangeValues1.getDisplayValues();
  var values2 = rangeValues2.getDisplayValues();
 
  var getSpreadsheet = function() {return rangeValues1};
  retry_(getSpreadsheet, numRetries);

  var dataHeader1 = values1[0];
  var dataValues1 = values1.slice(1).filter(row => {
      const dateIdx = dataHeader1.findIndex(val => val === 'Date');
      const date = moment(row[dateIdx], 'M/D/YYYY', true);
      return date.isValid() && date.isBetween(yearStart, yearEnd, 'day', '[]');
  });

  var dataHeader2 = values2[0];
  var dataValues2 = values2.slice(1).filter(row => {
      const dateIdx = dataHeader2.findIndex(val => val === 'RFQ Date');
      const date = moment(row[dateIdx], 'M/D/YYYY', true);
      return date.isValid() && date.isBetween(yearStart, yearEnd, 'day', '[]');
  });

  Logger.log(dataValues1.length)
  Logger.log(dataValues2.length)
 
  return {headerKR:dataHeader1, valuesKR:dataValues1, headerOS:dataHeader2, valuesOS:dataValues2};
}

function getRFQOverview() {

    var data_all = getFilteredRFQ();

    // Shared constants
    const periods = ['Weekly', 'Monthly', 'Quarterly', 'Yearly'];
    const statuses = ['Bidding', 'Pending', 'Ordered', 'Pass', 'Failed', 'Delete', 'Unselected', 'Overall'];
    const periodChecks = {
        Weekly: [weekStart, weekEnd],
        Monthly: [monthStart, monthEnd],
        Quarterly: [quarterStart, quarterEnd],
        Yearly: [yearStart, yearEnd],
    };

    const createPeriodObject = () =>
        Object.fromEntries(periods.map(period => [period, Object.fromEntries(statuses.map(status => [status, 0]))]));

    // Helper function to process data
    const processRFQData = (headerValues, bodyValues, dateCol, ownerCol, statusCol, targetUser) => {
        const final_obj = {
            Your: createPeriodObject(),
            Total: createPeriodObject()
        };

        // const headerValues = values[0];
        // const bodyValues = values.slice(1);

        // Get indices
        const dateIdx = headerValues.findIndex(val => val === dateCol);
        const ownerIdx = headerValues.findIndex(val => val === ownerCol);
        const statusIdx = headerValues.findIndex(val => val === statusCol);

        // Helper to increment counts
        const incrementCounts = (obj, period, status) => {
            obj[period][status]++;
            obj[period]['Overall']++;
        };

        // Filter duplicates (if necessary)
        const filteredValues = bodyValues.filter((arr, idx) =>
            bodyValues.findIndex(val => val[0] === arr[0]) === idx
        );

        // Process data
        filteredValues.forEach(arr => {
            const status = arr[statusIdx] || "Unselected";// || "Unselected"
            if (status !== 'Terminate') { //status && status !== 'Delete' && 
                const date = moment(arr[dateIdx], 'M/D/YYYY', true);
                if (date.isValid()) {
                    periods.forEach(period => {
                        const [start, end] = periodChecks[period];
                        if (date.isBetween(start, end, 'day', '[]')) {
                            if (arr[ownerIdx] === targetUser) {
                                incrementCounts(final_obj['Your'], period, status);
                            }
                            incrementCounts(final_obj['Total'], period, status);
                        }
                    });
                } else {
                    Logger.log(`Invalid date encountered: ${arr[dateIdx]}`);
                }
            }
        });

        return final_obj;
    };

    // Process KR and OS data
    const krHeader = data_all.headerKR;
    const osHeader = data_all.headerOS;
    const krValues = data_all.valuesKR;
    const osValues = data_all.valuesOS;
    // const krValues = getRFQ();
    // const osValues = getRFQOS().dataArray;
    const targetUser = getActiveUserName();
    //const targetUser = 'Y';

    const kr_obj = processRFQData(krHeader, krValues, 'Date', 'Owner', 'Status', targetUser);
    const os_obj = processRFQData(osHeader, osValues, 'RFQ Date', 'Owner', 'Status', targetUser);

    // Build range object
    const range_obj = Object.fromEntries(
        periods.map(period => {
            const [start, end] = periodChecks[period];
            return [period, `${start} ~ ${end}`];
        })
    );

    Logger.log({ KR: kr_obj, OS: os_obj, Range: range_obj });

    return { KR: kr_obj, OS: os_obj, Range: range_obj };
}


function getVendors(){
  var spreadSheetId = "1pWFTXnpKhO6VOzHT7KsjrUPxmKFwfQdlatTKmWvTjUk"; //CHANGE
  // var dataRange     = "list!A:XX"; //CHANGE
 
  // var range   = Sheets.Spreadsheets.Values.get(spreadSheetId,dataRange);
  
  // var values  = range.values;

  var sheetName1 = "Vendor List";
  //var sheetName2 = "pp/sop";

  var doc = SpreadsheetApp.openById(spreadSheetId);

  var sheet1 = doc.getSheetByName(sheetName1);
  //var sheet2 = doc.getSheetByName(sheetName2);

  var rangeValues1 = sheet1.getDataRange();
  //var values = rangeValues.getValues();
  var values1 = rangeValues1.getDisplayValues();
  //var rangeValues2 = sheet2.getDataRange();
  //var values2 = rangeValues2.getDisplayValues();

  var filledValues = values1.filter(arr => !arr.every(e => e === ""));

  Logger.log(filledValues.length)

  //values2.slice(1).forEach(arr => filledValues.push(arr));

  Logger.log(filledValues[2484])
  
  var getSpreadsheet = function() {return filledValues};
  retry_(getSpreadsheet, numRetries);
 
  return filledValues;
}


function getGmailEmails() {
  var userId = Session.getActiveUser().getEmail(); // Please modify this, if you want to use other userId.
  var threadList = Gmail.Users.Threads.list(userId).threads;
  
  var data_messages_arr = [];

  const getUrlPartsMessages = () => {
    const metadata = ['Subject', 'From', 'To'].map((key) => `metadataHeaders=${key}`).join('&');
    const data = {
      fields: 'messages(payload/headers)',
      //fields: 'payload/headers',
      format: `metadata`
    };
    const fields = Object.entries(data)
      .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
      .join('&');
    return `${fields}&${metadata}`;
  };
  const urlPartsMessages = getUrlPartsMessages();
  while(threadList.length){
    var body = threadList.splice(0,20).map(function(e){
        return {
            method: "GET",
            endpoint: "https://www.googleapis.com/gmail/v1/users/" + userId + "/threads/" + e.id + "?" + urlPartsMessages
        }
    }).filter(v => v !== undefined);
    //Logger.log(body)
    var url = "https://www.googleapis.com/batch/gmail/v1";
    //var boundary = "xxxxxxxxxx";
    var boundary = "batch_foobarbaz";
    var contentId = 0;
    var data = "--" + boundary + "\r\n";
    for (var i in body) {
      data += "Content-Type: application/http\r\n";
      data += "Content-ID: " + ++contentId + "\r\n\r\n";
      data += body[i].method + " " + body[i].endpoint + "\r\n\r\n";
      data += "--" + boundary + "\r\n";
    }
    var payload = Utilities.newBlob(data).getBytes();
    var options = {
      method: "post",
      contentType: "multipart/mixed; boundary=" + boundary,
      payload: payload,
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    };

    var res = UrlFetchApp.fetch(url, options).getContentText();
    var dat = res.split("--batch");
    var result = dat.slice(1, dat.length - 1).map(function(e){return e.match(/{[\S\s]+}/g)[0]});
    var threads = JSON.parse(JSON.stringify(result));
    var filteredThreads = threads.filter(val => val.includes('payload'));
    //var threadsError = threads.filter(val => val.includes('error'));
    Logger.log(threads)
    //Logger.log(threadsError.length

    filteredThreads.forEach((thread) => {
      var data_messages = JSON.parse(thread);
      var messageCount = data_messages.messages.length;
      //Logger.log(data_messages)

      data_messages.messages[messageCount - 1].payload.headers.forEach((header) => {
      //data_messages.payload.headers.forEach((header) => {
        if(header.name == "Subject"){
          subject = header.value.replaceAll(/Re: |Fwd: |FW: |\[FW\]|\[RE\]/ig,'');
        }
        if(header.name == "From"){
          if(header.value.includes('@')){
            
            sender = header.value.split(" <")[0].replaceAll(/"|'|<|>| *\([^)]*\) */ig,'');
            if(/(?=.*[A-Za-z])(?=.*[ㄱ-ㅣ가-힣]).*/.test(sender)) { sender = sender.replaceAll(/[^ㄱ-ㅣ가-힣]*/ig,''); }
            if(/^[ㄱ-ㅣ가-힣 ]+$/.test(sender)) { sender = sender.replaceAll(' ',''); }
            sender_client = header.value.includes(' <') ? header.value.split(" <")[1].split("@")[1].split(".")[0] : header.value.split("@")[1].split(".")[0];
          }
        }
        if(header.name == "To"){
          if(header.value.includes('@')){
            
            receiver = header.value.split(" <")[0].replaceAll(/"|'|<|>| via sales-kr| *\([^)]*\) */ig,'');
            if(/(?=.*[A-Za-z])(?=.*[ㄱ-ㅣ가-힣]).*/.test(receiver)) { receiver = receiver.replaceAll(/[^ㄱ-ㅣ가-힣]*/ig,''); }
            if(/^[ㄱ-ㅣ가-힣 ]+$/.test(receiver)) { receiver = receiver.replaceAll(' ',''); }
            receiver_client = header.value.includes(' <') ? header.value.split(" <")[1].split("@")[1].split(".")[0] : header.value.split("@")[1].split(".")[0];
          }
        }
      });

        
          var sender_check = sender.includes('@') ? sender.split("@")[0] : sender;
          data_messages_arr.push({ subject:subject, sender:sender_check, client:sender_client });
        
      
    });
  }
  //Logger.log(Session.getActiveUser())
  Logger.log(data_messages_arr.length)
  Logger.log(data_messages_arr)

  return data_messages_arr;
}

function getActiveUserEmail() {
  Logger.log(Session.getActiveUser().getEmail().split("@")[0]);
  return Session.getActiveUser().getEmail().split("@")[0];
}

function getActiveUserName() {
  const aboutData = Drive.About.get({
      fields: "user,storageQuota,importFormats,exportFormats"
    });
  const userName = aboutData["user"]["displayName"];
  const userEmail = aboutData["user"]["emailAddress"];
  return userName ? userName : userEmail.split("@")[0];
  //Logger.log(emailToOwner[Session.getActiveUser().getEmail().split("@")[0]]);
  //return emailToOwner[Session.getActiveUser().getEmail().split("@")[0]];
}

// 현재 QQ 페이지의 URL을 반환 | RETURNS CURRENT URL
// function getScriptURL() {
//   return ScriptApp.getService().getUrl();
// }
function getScriptURL(qs = null) {
  var url = ScriptApp.getService().getUrl();
  if(qs){
    if (qs.indexOf("?") === -1) {
      qs = "?" + qs;
    }
    url = url + qs;
  }
  return url;
}
 
// Index 파일에 JS & CSS 파일을 포함 시킴 | INCLUDE JAVASCRIPT AND CSS FILES
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// 함수: retry_
// 목적: 지정된 횟수만큼 함수를 재시도하고 실패 시 예외를 처리
// Parameters:
//   - func: 재시도할 함수
//   - numRetries: 재시도 횟수
//   - optLoggerFunction: 로그 기능을 제공하는 선택적 함수
function retry_(func, numRetries, optLoggerFunction) {
   // Ensure the number of retries is valid
  if (numRetries < 1) {
    throw new Error("Invalid number of retries: " + numRetries);
  }
  
  // Use default logger if none is provided
  optLoggerFunction = optLoggerFunction || function(msg) { Logger.log(msg); };

  for (var n = 0; n < numRetries; n++) {
    try {
      // 함수 실행 시도
      return func();
    } catch (e) {
      // 실패한 경우 예외 처리
      // if (optLoggerFunction) {
      //   optLoggerFunction("GASRetry " + n + ": " + e);
      // }
      // Log the error if a logger function is provided
      optLoggerFunction("Retry attempt " + (n + 1) + " failed: " + e.toString());
      
      if (n == (numRetries - 1)) {
        // 마지막 재시도에서도 실패하면 예외를 던짐
        //throw JSON.stringify(catchToObject_(e));
        // On the last retry, throw the error with more context
        throw new Error("Function failed after " + numRetries + " retries: " + e.toString());
      }

      // Exponential backoff with random jitter
      var waitTime = (Math.pow(2, n) * 1000) + Math.round(Math.random() * 1000);
      Utilities.sleep(waitTime);
    }    
  }
}

// 함수: spliceIntoChunks
// 목적: 배열을 지정된 크기의 청크로 나누어 반환
// Parameters:
//   - arr: 나눌 배열
//   - chunkSize: 청크 크기
function spliceIntoChunks(arr, chunkSize) {
    const res = [];
    while (arr.length > 0) {
        // 배열을 지정된 크기의 청크로 나누기
        const chunk = arr.splice(0, chunkSize);
        res.push(chunk);
    }
    return res;
}