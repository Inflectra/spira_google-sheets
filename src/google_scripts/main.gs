
//stubbed
var stubUser = {
    url: 'https://demo.spiraservice.net/christopher-abramson',
    userName: 'administrator',
    api_key: '&api-key=' + encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}')
  }

//App script boilerplate install function
//opens app on install
function onInstall(e) {
  onOpen(e);
}

//App script boilerplate open function
//opens sidebar
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}



//side bar function gets index.html and opens in side window
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Inflectra Corporation');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getProjects(currentUser){
  var fetcherURL = '/services/v5_0/RestService.svc/projects?username=';
  return fetcher(currentUser, fetcherURL);
}

function getUsers(currentUser, proj){
  var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/users?username=';
  return fetcher(currentUser, fetcherURL);
}

function getCustoms(currentUser, proj, artifact){
  var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/custom-properties/' + artifact + '?username=';
  return fetcher(currentUser, fetcherURL);
}

function fetcher (currentUser, fetcherURL, init){
  var URL = stubUser.url + fetcherURL + stubUser.userName + stubUser.api_key;
  var init = init || {'content-type' : 'application/json'}

  var response = UrlFetchApp.fetch(URL, init);

  return JSON.parse(response);
}


function error(type){
  if(type == 'impExp') {
    okWarn('There was an input error. Please check that your entries are correct.');
  } else if (type == 'unk') {
    okWarn('Unkown error. Please try again later or contact your system administrator');
  } else {
    okWarn('Network error. Please check your username, url, and password.');
  }
}

function success(string){
  // Show a 2-second popup with the title "Status" and the message "Task started".
  SpreadsheetApp.getActiveSpreadsheet().toast(string, 'Success', 2);
}


//Alert pop up for data clear warning
function warn(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will erase all unsaved changes. Continue?', ui.ButtonSet.YES_NO);

  //returns with user choice
  if (response == ui.Button.YES) {
    return true;
  } else {
    return false;
  }
}

//Alert pop up for project or artifact dropdown change
function warnProjArt(){
  okWarn('Warning! Changing the current project or artifact will clear all unsaved data.');
}

//Alert pop up for export success
function exportSuccess(){
  okWarn('Export Success! Clear sheet to export more artifacts.');
}

//Alert pop up for no template present
function noTemplate(){
  okWarn('Please load a template to continue.');
}

//warn with Ok button
function okWarn(dialoge){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(dialoge, ui.ButtonSet.OK);
}

//save function
function save(){
  //pop up telling the user that their data will be saved
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will save the current sheet in a new tab. Continue?', ui.ButtonSet.YES_NO);

  //returns with user choice
  if (response == ui.Button.YES) {
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getSheets()[0];
    //get entire spreadsheet id
    var id = spreadSheet.getId();
    //set as destination
    var destination = SpreadsheetApp.openById(id);
    //copy to destination
    sheet.copyTo(destination);
  }
}

//clear function
//clears current sheet
function clearAll(){
  //get first active spreadsheet
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheets()[0];

  //clear all formatting and content
  sheet.clear();
  //clears data validations from the entire sheet
  var range = SpreadsheetApp.getActive().getRange('A:AZ');
  range.clearDataValidations();
  //Reset sheet name
  sheet.setName('Sheet');
}


