


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


//ToDo
//onEdit function for text field coloration feature
//function onEdit(e){
//  // Set a comment on the edited cell to indicate when it was changed.
//  var range = e.range;
//  range.setNote('Last modified: ' + new Date());
//  Logger.log('changed')
//}


function getUsers(){
  //mock values
  var url = 'https://demo.spiraservice.net/christopher-abramson';
  var userName = 'administrator';
  var api_key = encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}');

  //build URL
  var usersURL = url
       + '/services/v5_0/RestService.svc/users/all?username='
       + userName
       +'&api-key='
       + api_key;

  var init = {'content-type': 'application/json'}

  var response = UrlFetchApp.fetch(usersURL, init)

  var data = JSON.parse(response);

  Logger.log(data)

  return data
}

function getUserData(url, userName, api_key){
  if(url && userName && api_key){
    model.currentUser = {'url': url, 'userName': userName, 'api-key': encodeURIComponent(api_key) }
  }
  //mock values
  var url = 'https://demo.spiraservice.net/christopher-abramson';
  var userName = 'administrator';
  var api_key = encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}');

  //build URL
  var projectsURL = url
       + '/services/v5_0/RestService.svc/projects?username='
       + userName
       +'&api-key='
       + api_key;

  var init = {'content-type': 'application/json'}

  var response = UrlFetchApp.fetch(projectsURL, init)

  var data = JSON.parse(response);
  //Logger.log(data.length);


  var projects = [];
  for(var i = 0; i < data.length; i++){
    projects.push(data[i].Name);
  };

  SpreadsheetApp.getActiveSpreadsheet().toast('Projects Loaded', 'Status', 2);

  return projects;
}



function loadTemplate(name, users){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  //set sheet name to model name
  sheet.setName(name)

  authors = { i: users }
  owners = { j : users }

  model.requirements.dropdowns.push(authors)
  model.requirements.dropdowns.push(owners)

  //color heading cells
  var stdColorRange = sheet.getRange('A1:M2');
  stdColorRange.setBackground('#ffbf80');
  var cusColorRange = sheet.getRange('N1:AQ2');
  cusColorRange.setBackground('#70db70');

  //set major headings on sheet and center
  sheet.getRange('A1:M1').merge().setValue("Requirements Standard Fields").setHorizontalAlignment("center");;
  sheet.getRange('N1:AQ1').merge().setValue("Custom Fields").setHorizontalAlignment("center");;

  //shorten varable name
  var store = model.requirements;

  //append headings to sheet
  sheet.appendRow(store.headings)

  //loop through model sizes data and set columns to correct width
  for(var i = 0; i < store.sizes.length; i++){
    sheet.setColumnWidth(store.sizes[i][0],store.sizes[i][1]);
  }

  //loop through dropdowns model data
  for(var i = 0; i < store.dropdowns.length; i++){
    //get column letter (as array)
    var column = Object.keys(store.dropdowns[i])
    var column = column[0];
    Logger.log(column);
    //get props
    var list = store.dropdowns[i][column]

    //set range to entire column excluding top two rows
    var cell = SpreadsheetApp.getActive().getRange(column + ':' + column).offset(2, 0);
    //require list of values as dropdown and entered values
    //require value in list: list variable is from the model, true shows dropdown arrow
    //allow invalid set to false does not allow invalid entries
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
    cell.setDataValidation(rule);

  }
}

function importer(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];

  var url = 'https://demo.spiraservice.net/christopher-abramson';
  var userName = 'administrator';
  var api_key = encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}');

  var param = 'projects/1/requirements?starting_row=1&number_of_rows=35&username=';
  var count = 'projects/1/requirements/count?username=';

  //build URL
  var requirementsURL = url
       + '/services/v5_0/RestService.svc/'
       + param
       + userName
       +'&api-key='
       + api_key;

  var countURL = url
       + '/services/v5_0/RestService.svc/'
       + count
       + userName
       +'&api-key='
       + api_key;


  var init = {'content-type': 'application/json'}

  var response = UrlFetchApp.fetch(requirementsURL, init)
  var data = JSON.parse(response);

  var range = sheet.getRange("A3:AQ3");

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
    range.getCell(ss_i, 9).setValue(data[i].AuthorId);
    range.getCell(ss_i, 10).setValue(data[i].OwnerId);
    range.getCell(ss_i, 11).setValue(data[i].ComponentId);

    range = range.offset(1, 0, 43);
 }
}



//Bloated export function to be refactored

function exporter(proj, list){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];
  var projectNum = '';

  for(var i = 0; i < list.length; i++){
    if(proj === list[i]){
      projectNum = i;
    }
  }

  var range = sheet.getRange("A3:AQ3")
  var isRangeEmpty = false;

  var numberOfRows = 0;
  var count = 0;
  var bodyArr = [];

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

    //loop through cells in row
    for (var i = 0; i < model.requirements.JSON_headings.length; i++){

      //get cell value
      var cell = range.offset(j, i).getValue();

      if (i === 5) {
        switch (cell){
          case '1 - Critical':
            cell = 1
            break;
          case '2 - High':
            cell = 2
            break;
          case '3 - Medium':
            cell = 3
            break;
          case '4 - Low':
            cell = 4
            break;
          default:
            cell = null;
        }
      }

      //get indent amount
      var indentCount = 0;
      //check for indent character '>'
      if(cell && cell[0] === '>'){
        //increment indent counter while there are '>'s present
        while (cell[0] === '>'){
          //get entry length for slice
          var len = cell.length;
          //slice the first character off of the entry
          cell = cell.slice(1, len);
          indentCount++;
        }
      }

      //get cell font weight
      var cellBold = range.offset(j, i).getFontWeight();
      //check to see if the cell is bold
      //if true wrap in bold tags
      cellBold === 'bold' ? cell = '<b>' + cell + '</b>' : null

      //get cell font style
      var cellItalic = range.offset(j, i).getFontStyle();
      //check to see if the cell is italic
      //if true wrap in em tags
      cellItalic === 'italic' ? cell = '<em>' + cell + '</em>' : null

      //if empty add null other wise add the cell
      // ...to the object under the proper key relative to its location on the template
      if (cell === ""){
        xObj[model.requirements.JSON_headings[i]] = null;
      } else {
        xObj[model.requirements.JSON_headings[i]] = cell;
      }

      //todo: something with indentCount
    }
    //if not empty add object
    if (xObj !== {}) {
      bodyArr.push(xObj)
    }
  }


  //return exportCall( bodyArr[0] )
  return JSON.stringify( bodyArr[0] )
  //return bodyArr[0];
}

function exportCall(body){
    //stub
  var url = 'https://demo.spiraservice.net/christopher-abramson';
  var userName = 'administrator';
  var api_key = encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}');
  var projectNum = 1;

  var exportURL = url
       + '/services/v5_0/RestService.svc/projects/'
       + projectNum
       + '/requirements?username='
       + userName
       +'&api-key='
       + api_key;

  var init = {
   'method' : 'post',
   'contentType': 'application/json',
   'payload' : body
  };

  var response = UrlFetchApp.fetch(exportURL, init)
  //var data = JSON.parse(response);

  return response;

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





//var API_KEY = encodeURIComponent('{2AE93998-6849-4132-80F6-3C9981A7CB96}');
//var USER_NAME = 'administrator';
//var URL_END = '&username=' + USER_NAME + '&api-key=' + API_KEY;
//var URL_BEG = '/services/v5_0/RestService.svc/';
//var param = 'projects/1/requirements?starting_row=1&number_of_rows=35';
//var count = 'projects/1/requirements/count';
//var final_URL = URL_BEG + param + URL_END;
//function tester(){}




//create templates for Test Cases
//make it so the template is selected by the drop down

//export (data in)

    //import data



//authentication
//remember user


//number of entries in artifact

//google style guide
//logos





//options menu?

//UX
//get rid of connect button

//options field

//add html warning page thats hidden with function that returns true or false for yes or no






