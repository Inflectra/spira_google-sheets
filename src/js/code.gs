
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

function getUsers(currentUser){
  var params = '/services/v5_0/RestService.svc/users/all?username='
  var res = fetcher(currentUser, params);

  return res;
}

function getCustoms(currentUser){
  var params = '/services/v5_0/RestService.svc/requirements?username='
  var res = fetcher(currentUser, params);
  Logger.log(res)
  return res;
}

function fetcher (currentUser, params, init){
  var URL = stubUser.url + params + stubUser.userName + stubUser.api_key;
  init ? null : init = {'content-type' : 'application/json'}


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
  var stdColorRange = sheet.getRange('A1:M2');
  stdColorRange.setBackground('#ffbf80');
  var cusColorRange = sheet.getRange('N1:AQ2');
  cusColorRange.setBackground('#70db70');
  var reqIdRange = sheet.getRange('A3:A100');
  reqIdRange.setBackground('#a6a6a6')

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
    var list = data.requirements.dropdowns[name]


    //set range to entire column excluding top two rows
    var cell = SpreadsheetApp.getActive().getRange(letter + ':' + letter).offset(2, 0);
    //require list of values as dropdown and entered values
    //require value in list: list variable is from the model, true shows dropdown arrow
    //allow invalid set to false does not allow invalid entries
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
    cell.setDataValidation(rule);
  }
}

function customFieldSetter(range, data){
  //shorten variable
  var fields = data.requirements.customFields
  //loop through model custom fields data
  //take passed in range and only overwrite the fields if a value is present in the model
  for(var i = 0; i < fields.length; i++){
    var cell = range.getCell(1, i + 1)
    cell.setValue('Custom Field ' + (i + 1) + '\n' + fields[i].Definition.Name).setWrap(true);

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

function mapper(item, list, objNums){
  var val = 1;
  if(objNums){
    for (var i = 1; i < list.length; i++){
      if (item == list[i][0]) {val = list[i][1]}
    }
  } else {
    for (var i = 0; i < list.length; i++){
      if (item == list[i]){ val = i }
    }
  }
  //Logger.log(list)
  Logger.log(item)
  return val;
}

//function richData(data){
//  var textArr = data.split(' ');
//
//  for (var i = 0; i < textArr.length; i++){
//
//    var word = textArr[i]
//    //.isBold()
//
////    var italic = textArr[i].getFontStyle()
////    var underline = textArr[i].getFOntLines()
////
//    Logger.log(word)// + italic + underline
//
//  }
//  return data;
//}

function indender(cell){
  // var indentCount = 0;
  // //check for indent character '>'
  // if(cell && cell[0] === '>'){
  // //increment indent counter while there are '>'s present
  //   while (cell[0] === '>'){
  //     //get entry length for slice
  //     var len = cell.length;
  //     //slice the first character off of the entry
  //     cell = cell.slice(1, len);
  //     indentCount++;
  //   }
  //   xObj['IndentLevel'] = 'AAB';
  // }
  return 'AAA'
}

function exporter(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];

  var range = sheet.getRange("A3:AQ3")
  var isRangeEmpty = false;
  var numberOfRows = 0;
  var count = 0;
  var bodyArr = [];

  //loop through and collect number of rows that contain data
  //TODO skip two lines before changing isRangeEmpty var
  while (isRangeEmpty === false){
    var newRange = range.offset(count, 0, 43);
    if ( newRange.isBlank() ){
      isRangeEmpty = true
    } else {
      //move to next row
      count++;
      //add to number of rows
      numberOfRows++;
    }
  }

  //loop through rows
  for (var j = 0; j < numberOfRows + 1; j++){

    //initialize/clear new object for row values
    var xObj = {}
    //shorten variable
    var reqs = data.templateData.requirements;

    //loop through cells in row
    for (var i = 0; i < reqs.JSON_headings.length; i++){

      //get cell value
      var cell = range.offset(j, i).getValue();

      //passes description data to richData function to attach HTML tags for spirateam
      //if(i === 2.0){ cell = richData(cell) }

      //shorten variables
      var users = data.userData.projUserWNum;
      var dataReqs = data.templateData.requirements;

      //pass values to mapper function
      //mapper iterates and assigns the values number based on the list order
      if(i === 4.0){ cell = mapper(cell, dataReqs.dropdowns['Type']) }

      if(i === 5.0){ xObj['ImportanceId'] = mapper(cell, dataReqs.dropdowns['Importance']) }

      if(i === 6.0){ xObj['StatusId'] = mapper(cell, dataReqs.dropdowns['Status']) }

      if (i === 8.0){ xObj['AuthorId'] = mapper(cell, users, true) }

      if (i === 9.0){ xObj['OwnerId'] = mapper(cell, users, true) }



      //call indent checker and set indent amount
      xObj['IndentLevel'] = indender();

      //if empty add null otherwise add the cell
      // ...to the object under the proper key relative to its location on the template
      //Offset by 2 for proj name and indent level
      if (cell === ""){
        xObj[reqs.JSON_headings[i]] = null;
      } else {
        xObj[reqs.JSON_headings[i]] = cell;
      }

    }

    //if not empty add object or a generated placeholder (no name)
    if ( xObj.Name ) {
      xObj['ProjectName'] = data.templateData.currentProjectName;
      bodyArr.push(xObj)
    }

  }

  // set up to individually add each requirement to spirateam
  // maybe there's a way to bulk add them instead of individual calls?
  var responses = []
  for(var i = 0; i < bodyArr.length; i++){
   //stringify
   var JSON_body = JSON.stringify( bodyArr[i] );
   //send JSON to export function
   var response = requirementExportCall( JSON_body, data.templateData.currentProjectNumber, data.userData.currentUser )
   //push API approval into array
   responses.push(response.RequirementId)
  }




  return responses
  //return bodyArr
  //return JSON.stringify( bodyArr )
  //return JSON_body;
}

function requirementExportCall(body, projNum, currentUser){
  //unique url for requirement POST
  var params = '/services/v5_0/RestService.svc/projects/' + projNum + '/requirements?username=';
  //POST headers

  var init = {
   'method' : 'post',
   'contentType': 'application/json',
   'payload' : body
  };
  //call fetch with POST request
  var res = fetcher(currentUser, params, init);

  return res;
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