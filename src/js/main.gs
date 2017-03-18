
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
  Logger.log(filename)
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


function templateLoader(data){
  clearAll();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var dropdownColumnAssignments = [
        ['Type', 'e'],['Importance', 'f'], ['Status', 'g'], ['Author', 'i'], ['Owner', 'j']
      ]

  //set sheet name to model name
  sheet.setName(data.currentProjectName + ' - ' + data.currentArtifactName);

  //color heading cells
  var stdColorRange = sheet.getRange(data.requirements.standardRange);
  stdColorRange.setBackground('#ffbf80');
  var cusColorRange = sheet.getRange(data.requirements.customRange);
  cusColorRange.setBackground('#70db70');
  var reqIdRange = sheet.getRange('A3:A100');
  reqIdRange.setBackground('#a6a6a6')
  //set column A to present a warning if the user trys to write in a value
  var protection = reqIdRange.protect().setDescription('Exported items must not have a requirement number');
  //set warning. Remove this to make the column un-writable
  protection.setWarningOnly(true);

  sheet.getRange('A1:M1').merge().setValue("Requirements Standard Fields").setHorizontalAlignment("center");
  sheet.getRange('N1:AQ1').merge().setValue("Custom Fields").setHorizontalAlignment("center");

  //append headings to sheet
  sheet.appendRow(data.requirements.headings)

  //set custom headings if they exist
  //pass in custom field range and data model
  customFieldSetter(sheet.getRange('N2:AQ2'), data);

  //loop through model sizes data and set columns to correct width
  for(var i = 0; i < data.requirements.sizes.length; i++){
    sheet.setColumnWidth(data.requirements.sizes[i][0],data.requirements.sizes[i][1]);
  }

  //loop through dropdowns model data
  for(var i = 0; i < dropdownColumnAssignments.length; i++){
    var letter = dropdownColumnAssignments[i][1];
    var name = dropdownColumnAssignments[i][0];
    var list = [];
    if (name == 'Owner' || name == 'Author'){
      list = data.requirements.dropdowns[name]
    } else {
      var listArr = [];
      //loop through 2D arrays and form standard array
      for(var j = 0; j < data.requirements.dropdowns[name].length; j++){
        listArr.push(data.requirements.dropdowns[name][j][1])
      }
      //list must be an array so assign new array to list variable
      list = listArr;
    }

    //set range to entire column excluding top two rows
    var cell = SpreadsheetApp.getActive().getRange(letter + ':' + letter).offset(2, 0);
    //require list of values as dropdown and entered values
    //require value in list: list variable is from the model, true shows dropdown arrow
    //allow invalid set to false does not allow invalid entries
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
    cell.setDataValidation(rule);
  }
  //set number only columns to only accept numbers
  for(var i = 0; i < data.requirements.requireNumberFields.length; i++){
    var colLetter = data.requirements.requireNumberFields[i];
    var column = SpreadsheetApp.getActive().getRange(colLetter + ':' + colLetter);
    var rule = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).setAllowInvalid(false).setHelpText('Must be a positive integer').build();
    column.setDataValidation(rule);
  }
}

function customFieldSetter(range, data){
  //shorten variable
  var fields = data.requirements.customFields
  //loop through model custom fields data
  //take passed in range and only overwrite the fields if a value is present in the model
  for(var i = 0; i < fields.length; i++){
    var cell = range.getCell(1, i + 1)
    cell.setValue('Custom Field ' + (i + 1) + '\n' + fields[i].Name).setWrap(true);

  }
}



//import function, basic for now
//stretch goal is to have this as a useful import
function importer(currentUser){
  //get spreadsheet and active first sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];

  // needed for eventual actual importer
  // var paramsCount = '/services/v5_0/RestService.svc/projects/1/requirements/count?username=';
  // var count = getFetch(currentUser, paramsCount )

  //call defined fetch function
  //current params has count set to 35, this can be set/changed programmatically with the count call listed above (stretch goal)
  var params = '/services/v5_0/RestService.svc/projects/1/requirements?starting_row=1&number_of_rows=35&username=';
  var data = fetcher(currentUser, params)

  //get first row range
  var range = sheet.getRange("A3:AQ3");

  //loop through cells in range
  for(var i = 0; i < data.length; i++){
    var ss_i = i + 1
    range.getCell(ss_i, 1).setValue(data[i].RequirementId);
    range.getCell(ss_i, 2).setValue(data[i].Name);
    range.getCell(ss_i, 3).setValue(data[i].Description);
    range.getCell(ss_i, 4).setValue(data[i].ReleaseVersionNumber);
    range.getCell(ss_i, 5).setValue(data[i].RequirementTypeName);
    range.getCell(ss_i, 6).setValue(data[i].ImportanceName);
    range.getCell(ss_i, 7).setValue(data[i].StatusName);
    range.getCell(ss_i, 8).setValue(data[i].EstimatePoints);
    range.getCell(ss_i, 9).setValue(data[i].AuthorName);
    range.getCell(ss_i, 10).setValue(data[i].OwnerName);
    range.getCell(ss_i, 11).setValue(data[i].ComponentId);

    //moves the range down one row
    range = range.offset(1, 0, 43);
 }
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