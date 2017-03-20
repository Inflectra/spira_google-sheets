
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
  var params = '/services/v5_0/RestService.svc/projects?username='
  var res = fetcher(currentUser, params);

  return res;
}

function getUsers(currentUser, proj){
  var params = '/services/v5_0/RestService.svc/projects/' + proj + '/users?username='
  var res = fetcher(currentUser, params);

  return res;
}

function getCustoms(currentUser, proj, artifact){
  var params = '/services/v5_0/RestService.svc/projects/' + proj + '/custom-properties/' + artifact + '?username='
  var res = fetcher(currentUser, params);

  return res;
}

function fetcher (currentUser, params, init){
  var URL = stubUser.url + params + stubUser.userName + stubUser.api_key;
  var init = init || {'content-type' : 'application/json'}

  var response = UrlFetchApp.fetch(URL, init)

  return JSON.parse(response);
}




function error(type){
  var ui = SpreadsheetApp.getUi();
  if(type == 'impExp') {
    var response = ui.alert('There was an input error. Please check that your entries are correct.', ui.ButtonSet.OK);
  } else if (type == 'unk') {
    var response = ui.alert('Unkown error. Please try again later or contact your system administrator', ui.ButtonSet.OK);
  } else {
    var response = ui.alert('Network error. Please check your username, url, and password.', ui.ButtonSet.OK);
  }
}

function success(string){
  // Show a 3-second popup with the title "Status" and the message "Task started".
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

function warnProjArt(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(' Warning! Changing the current project or artifact will clear all unsaved data.', ui.ButtonSet.OK);
}


//Alert pop up for no template present
function noTemplate() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Please load a template to continue.', ui.ButtonSet.OK);
}

//save function
function save(){
  //pop up telling the user that their data will be saved
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will save the current sheet in a new tab. Continue?', ui.ButtonSet.YES_NO);

  //returns with user choice
  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    //get entire spreadsheet id
    var id = ss.getId()
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