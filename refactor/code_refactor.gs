
//stubbed
var stubUser = {
    url: 'https://demo.spiraservice.net/christopher-abramson',
    userName: 'administrator',
    api_key: '&api-key=' + encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}')
  }

//App script boilerplate open function
//opens sidebar
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

//App script boilerplate install function
//opens app on install
function onInstall(e) {
  onOpen(e);
}

//side bar function gets index.html and opens in side window
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Inflectra Corporation');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getProjects(currentUser){
  var params = '/services/v5_0/RestService.svc/projects?username='
  var res = getFetch(currentUser, params);

  return res;
}

function getUsers(currentUser){
  var params = '/services/v5_0/RestService.svc/users/all?username='
  var res = getFetch(currentUser, params);

  return res;
}

function getFetch (currentUser, params){
  var URL = stubUser.url + params + stubUser.userName + stubUser.api_key;
  var init = {'content-type' : 'application/json'}

  var response = UrlFetchApp.fetch(URL, init)

  return JSON.parse(response);
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

function templateLoader(data){
  Logger.log(data);
}




//Alert pop up for no template present
function noTemplate() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Please load a template to continue.', ui.ButtonSet.OK);
}

//clear function
//clears current sheet
function clearAll(){
  //get first active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];

  //clear all formatting and content
  sheet.clear()
  //clears data validations from the entire sheet
  var range = SpreadsheetApp.getActive().getRange('A:AZ');
  range.clearDataValidations();
  //Reset sheet name
  sheet.setName('Sheet');
}