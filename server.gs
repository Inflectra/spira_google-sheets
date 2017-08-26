/*
 * =================
 * UTILITY FUNCTIONS
 * =================
 * 
 * The Utility functions needed for initialization 
 * and basic app functionality are located here, as well as all GET functions. 
 * All Google App Script (GAS) files are bundled by the engine 
 * at start up so any non-scoped variable declared will be available globally.
 *
 */

// App script boilerplate install function
// opens app on install
function onInstall(e) {
  onOpen(e);
}


// App script boilerplate open function
// opens sidebar
// Method `addItem`  is related to the 'Add-on' menu items. Currently just one is listed 'Start' in the dropdown menu
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Start', 'showSidebar').addToUi();
}


// side bar function gets index.html and opens in side window
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Inflectra Corporation');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}


// This function is part of the google template engine and allows for modularization of code
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



/*
**********************
Fetch `GET` functions
**********************
*/

//This function is called on initial log in and acts as user validation
//Gets projects for current logged in user and returns data to scripts.js.html
function getProjects(currentUser) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects?username=';
    return fetcher(currentUser, fetcherURL);
}

/*
All of these functions are called when a template is loaded and they return data to scripts.js.html

When new artifacts are added new GET functions will need to be added and removed.
*/

//Gets User data for current user and current project users
function getUsers(currentUser, proj) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/users?username=';
    return fetcher(currentUser, fetcherURL);
}

//Gets custom fields for current user, project and artifact.
function getCustoms(currentUser, proj, artifact) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/custom-properties/' + artifact + '?username=';
    return fetcher(currentUser, fetcherURL);
}

//Gets Releases for current user and project.
function getReleases(currentUser, proj) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/releases?username=';
    return fetcher(currentUser, fetcherURL);
}

//Gets components for current user and project.
function getComponents(currentUser, proj) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/components?active_only=true&include_deleted=false&username=';
    return fetcher(currentUser, fetcherURL);
}


//Fetch function uses Googles built in fetch api
//Arguments are current user object and url params
function fetcher(currentUser, fetcherURL) {

    //google base 64 encoded string utils
    var decoded = Utilities.base64Decode(currentUser.api_key);
    var APIKEY = Utilities.newBlob(decoded).getDataAsString();

    //build URL from args
    //this must be changed if using mock values in development
    var URL = currentUser.url + fetcherURL + currentUser.userName + APIKEY;
    //set MIME type
    var init = { 'content-type': 'application/json' };
    //call Google fetch function
    var response = UrlFetchApp.fetch(URL, init);
    //returns parsed JSON
    //unparsed response contains error codes if needed
    return JSON.parse(response);
}

/*
*************
Error Functions
*************
*/

//Error notification function
//Assigns string value and routes error call from scripts.js.html
//Argument `type` is a string identifying the message to be displayed
function error(type) {
    if (type == 'impExp') {
        okWarn('There was an input error. Please check that your entries are correct.');
    } else if (type == 'unk') {
        okWarn('Unkown error. Please try again later or contact your system administrator');
    } else {
        okWarn('Network error. Please check your username, url, and password. If correct make sure you have the correct permissions.');
    }
}

//Pop-up notification function
//Argument `string` is the message to be displayed
function success(string) {
    // Show a 2-second popup with the title "Status" and a message passed in as an argument.
    SpreadsheetApp.getActiveSpreadsheet().toast(string, 'Success', 2);
}


//Alert pop up for data clear warning
function warn(messageString) {
    var ui = SpreadsheetApp.getUi();
    //alert popup with yes and no button
    var response = ui.alert(messageString, ui.ButtonSet.YES_NO);

    //returns with user choice
    if (response == ui.Button.YES) {
        return true;
    } else {
        return false;
    }
}

//Alert pop up for export success
//Argument `err` is a boolean sent from the export function
function exportSuccess(err) {
    if (err) {
        okWarn('Operation complete, some errors occurred. Clear sheet to export more artifacts.');
    } else {
        okWarn('Operation complete. Clear sheet to export more artifacts.');
    }
}

//Alert pop up for no template present
function noTemplate() {
    okWarn('Please load a template to continue.');
}

//Google alert popup with Ok button
function okWarn(dialoge) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(dialoge, ui.ButtonSet.OK);
}


/*
************
Utilities
************
*/

//save function
function save() {
    //pop up telling the user that their data will be saved
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('This will save the current sheet in a new tab. Continue?', ui.ButtonSet.YES_NO);

    //returns with user choice
    if (response == ui.Button.YES) {
        //get first tab of  active spreadsheet
        var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = spreadSheet.getSheets()[0];

        //get entire open spreadsheet id
        var id = spreadSheet.getId();

        //set current spreadsheet file as destination
        var destination = SpreadsheetApp.openById(id);

        //copy tab to current spreadsheet in new tab
        sheet.copyTo(destination);
      
        //returns true to que success popup
        return true;
    } else {
        //returns false to ignore success popup     
        return false;
    }
}

//clear function
//clears current sheet
function clearAll() {
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
















/*
 * =================
 * TEMPLATE CREATION
 * =================
 * 
 * This function creates a template based on the model template data 
 * TODO: currently only creates requirements/task template in non generic way
 * Takes the entire data model as an argument
 *
 */

//function for template creation
function templateLoader(data) {
    //call clear function and clear spreadsheet depending on user input
    clearAll();

    //select open file and select first tab
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getSheets()[0];
    var artifactData = data[data.currentArtifactName];
    

    //shorten variable
    var dropdownColumnAssignments = artifactData.dropdownColumnAssignments;

    //set sheet (tab) name to model name
    sheet.setName(data.currentProjectName + ' - ' + data.currentArtifactName);

    //set heading colors and font colors for standard and custom ranges
    var stdColorRange = sheet.getRange(artifactData.standardRange);
    stdColorRange.setBackground('#073642');
    stdColorRange.setFontColor('#fff');

    var cusColorRange = sheet.getRange(artifactData.customRange);
    cusColorRange.setBackground('#1398b9');
    cusColorRange.setFontColor('#fff');

    //get range for artifact ids and set color
    //color set to grey to denote unwritable field
    var reqIdRange = sheet.getRange('A3:A400');
    reqIdRange.setBackground('#a6a6a6');

    //set customfield cells as grey if inactive
    var customCellRange = sheet.getRange('N3:AQ400');
    customCellRange.setBackground('#a6a6a6');

    //unsupported fields also colored grey
    for (var i = 0; i < artifactData.unsupported.length; i++) {
        var column = sheet.getRange(artifactData.unsupported[i]);
        column.setBackground('#a6a6a6')
    }

    //set column A to present a warning if the user tries to write in a value
    var protection = reqIdRange.protect().setDescription('Exported items must not have a requirement number');
    //set warning. Remove this to make the column un-writable
    protection.setWarningOnly(true);

    //set title range and center
    sheet.getRange(artifactData.standardTitleRange).merge().setValue("Standard Fields").setHorizontalAlignment("center");
    sheet.getRange(artifactData.customTitleRange).merge().setValue("Custom Fields").setHorizontalAlignment("center");

    //append standard column headings to sheet
    sheet.appendRow(artifactData.headings)

    //set custom headings if they exist
    //pass in custom field range, data model, and custom column to be used for background coloring
    customHeadSetter(sheet.getRange(artifactData.customHeaders), data, sheet.getRange(artifactData.customColumnLength));

    //loop through model size data and set columns to correct width
    for (var i = 0; i < artifactData.sizes.length; i++) {
        sheet.setColumnWidth(artifactData.sizes[i][0], artifactData.sizes[i][1]);
    }

    //main custom field function assigns type, dropdowns, datavalidation etc. See function for details.
    customContentSetter(sheet.getRange(artifactData.customCellRange), data)

    //loop through dropdowns model data
    for (var i = 0; i < dropdownColumnAssignments.length; i++) {
        //variable assignment from dropdown object
        var letter = dropdownColumnAssignments[i][1];
        var name = dropdownColumnAssignments[i][0];
        //array that will hold dropdown values
        var list = [];
        //loop through 2D arrays and form standard array
        for (var j = 0; j < artifactData.dropdowns[name].length; j++) {
            list.push(artifactData.dropdowns[name][j][1])

          
        }

        //set range to entire column excluding top two rows (offset)
        var cell = SpreadsheetApp.getActive().getRange(letter + ':' + letter).offset(2, 0);
        //require list of values as a dropdown
        //require value in list: list variable is from the model, true shows dropdown arrow
        //allow invalid set to false does not allow invalid entries
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
        cell.setDataValidation(rule);
    }
  
  
 
  
    //loop through data model
    //set 'number only' columns to only accept numbers
    for (var i = 0; i < artifactData.requireNumberFields.length; i++) {
        var colLetter = artifactData.requireNumberFields[i];
        var column = SpreadsheetApp.getActive().getRange(colLetter + ':' + colLetter);
        //does not allow negative numbers or non-integers
        //sets a tooltip explaining cell rules
        var rule = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).setAllowInvalid(false).setHelpText('Must be a positive integer').build();
        column.setDataValidation(rule);
    }
}

/*
Custom header setter function

Takes a range of cells, the data model and a column range as arguments
*/

//Sets headings for custom fields
function customHeadSetter(range, data, col) {

    //shorten variable
    var fields = data.requirements.customFields

    //loop through model custom fields data
    //take passed in range and only overwrite the fields if a value is present in the model
    for (var i = 0; i < fields.length; i++) {
        //get cell and offset by one column every iteration
        var cell = range.getCell(1, i + 1)
            //set heading and wrap text to fit
        cell.setValue('Custom Field ' + (i + 1) + '\n' + fields[i].Name).setWrap(true);
        //get column and offset (move to the right) every iteration and set background
        var column = col.offset(0, i)
        column.setBackground('#fff');
    }
}

/*
Custom content setter function

Sets the data validation rules for the custom fields

Takes a range of cells and the data model as arguments.
*/

//Sets dropdown and validation content for custom fields
function customContentSetter(range, data) {
    //shorten variable
    var customs = data.requirements.customFields;
    //grab users list from owners dropdown
    var users = data.requirements.dropdowns['Owner'];
    //loop through custom property fields
    for (var i = 0; i < customs.length; i++) {

        //check if field matches {2: integer} or {3: float}
        if (customs[i].CustomPropertyTypeId == 2 || customs[i].CustomPropertyTypeId == 3) {

            //get first cell in range
            var cell = range.getCell(1, i + 1);

            //get column range (x : x)
            //gets the column letter of the selected cell, i.e 'F'
            var column = columnRanger(cell);

            //set number only data validation
            //must be a valid number greater than -1 (also excludes 1.1.0 style numbers)
            var rule = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).setAllowInvalid(false).setHelpText('Must be a positive integer').build();
            column.setDataValidation(rule);
        }

        //check if field matches {4: boolean}
        if (customs[i].CustomPropertyTypeId == 4) {

            //dropdown options
            //'True' and 'False' don't work as dropdown choices
            var list = ["Yes", "No"];

            //get first cell in range
            var cell = range.getCell(1, i + 1);

            //get A1 notation from google range dataType
            var cellsTop = cell.getA1Notation();

            // set the end of the column
            //needed to apply data validation, I've set the column to be 200 cells long
            var cellsEnd = cell.offset(200, 0).getA1Notation();

            //sets the column in A1 notation (XX : XX)
            var column = SpreadsheetApp.getActive().getRange(cellsTop + ':' + cellsEnd);

            //builds the data validation rule
            var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
            column.setDataValidation(rule);
        }

        //check if field matches {5: date}
        if (customs[i].CustomPropertyTypeId == 5) {
            var cell = range.getCell(1, i + 1);

            //gets the column letter of the selected cell, i.e 'F'
            var column = columnRanger(cell);

            //set number only data validation
            var rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Must be a valid date').build();
            column.setDataValidation(rule);
        }

        //List {6} and MultiList {7}
        if (customs[i].CustomPropertyTypeId == 6 || customs[i].CustomPropertyTypeId == 7) {
            var list = [];
            //loop through the custom list values and push into our holder array
            for (var j = 0; j < customs[i].CustomList.Values.length; j++) {
                list.push(customs[i].CustomList.Values[j].Name);
            }
            //get the first cell in the column
            var cell = range.getCell(1, i + 1);

            //get the top and bottom of the range i.e (A1:A200)
            var cellsTop = cell.getA1Notation();
            var cellsEnd = cell.offset(200, 0).getA1Notation();
            var column = SpreadsheetApp.getActive().getRange(cellsTop + ':' + cellsEnd);

            //assign dropdowns and do not allow entries outside of the supplied list
            var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
            column.setDataValidation(rule);
        }

        //users
        if (customs[i].CustomPropertyTypeId == 8) {
            //loop through list of users and assign them to a holder array
            var list = [];
            var len = users.length;
            for (var j = 0; j < len; j++) {
                list.push(users[j][1]);
            }

            //get the top and bottom of the range i.e (A1:A200)
            var cell = range.getCell(1, i + 1);
            var cellsTop = cell.getA1Notation();
            var cellsEnd = cell.offset(200, 0).getA1Notation();
            var column = SpreadsheetApp.getActive().getRange(cellsTop + ':' + cellsEnd);

            //assign dropdowns and do not allow entries outside of the supplied list
            var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
            column.setDataValidation(rule);
        }
    }

}

//supplies the column of the current cell
function columnRanger(cell) {
    //get the column
    var col = cell.getColumn();
    //get column letter
    col = columnToLetter(col);
    //get column range for data validation
    var column = SpreadsheetApp.getActive().getRange(col + ':' + col);

    return column;
}

//open source column to letter function **Adam L from Stack OverFlow
function columnToLetter(column) {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}













/*
 * ================
 * SENDING TO SPIRA
 * ================
 * 
 * The main function takes the entire data model and the artifact type 
 * and calls the child function to set various object values before 
 * sending the finished objects to SpiraTeam
 *
 */

function exporter(data, artifactType) {
    //get the active spreadsheet and first tab
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getSheets()[0];

    //range of cells in a row for the given artifact
    var range = sheet.getRange(data.templateData.requirements.cellRange);
    //range of cells in a row for custom fields
    var customRange = sheet.getRange(data.templateData.requirements.customCellRange);
    var isRowEmpty = false;
    var numberOfRows = 0;
    var row = 0;
    var col = 0;

    //final arrays that hold finished objects for export
    var responses = [];
    var xObjArr = [];

    //shorten variable
    var reqs = data.templateData.requirements;

    //Model window
    var htmlOutput = HtmlService.createHtmlOutput('<p>Preparing your data for export!</p>').setWidth(250).setHeight(75);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Progress');

    //loop through and collect number of rows that contain data
    while (isRowEmpty === false) {
        //select row i.e (0, 0, 43)
        //the offset method moves the row down each iteration
        var newRange = range.offset(row, col, reqs.cellRangeLength);
        //check if the row is empty
        if (newRange.isBlank()) {
            //if row is empty set var to true
            isRowEmpty = true
        } else {
            //move to next row
            row++;
            //add to number of rows
            numberOfRows++;
        }
    }

    //loop through standard data rows
    for (var j = 0; j < numberOfRows + 1; j++) {

        //initialize/clear new object for row values
        var xObj = {}

        //send data model and current row to custom data function
        var row = customRange.offset(j, 0)
        xObj['CustomProperties'] = customHeaderRowBuilder(data, row)

        //set position number
        //used for indent
        xObj['positionNumber'] = 0;

        //loop through cells in row according to the JSON headings
        for (var i = 0; i < reqs.JSON_headings.length; i++) {

            //get cell value
            var cell = range.offset(j, i).getValue();

            //get cell Range for id number insertion after export
            if (i === 0.0) { xObj['idField'] = range.offset(j, i).getCell(1, 1) }

            //call indent checker and set indent amount
            if (i === 1.0) {
                //call indent function
                //counts the number of ">"s to assign an indent value
                xObj['indentCount'] = indenter(cell)

                //remove '>' symbols from requirement name string
                while (cell[0] == '>' || cell[0] == ' ') {
                    //removes first character if it's a space or ">"
                    cell = cell.slice(1)
                }
            }

            //shorten variables
            var users = data.userData.projUserWNum;

            //pass values to mapper function
            //mapper iterates and assigns the values number based on the list order
            if (i === 3.0) { xObj['ReleaseId'] = mapper(cell, reqs.dropdowns['Version Number']) }

            if (i === 4.0) { cell = mapper(cell, reqs.dropdowns['Type']) }

            if (i === 5.0) { xObj['ImportanceId'] = mapper(cell, reqs.dropdowns['Importance']) }

            if (i === 6.0) { xObj['StatusId'] = mapper(cell, reqs.dropdowns['Status']) }

            if (i === 8.0) { xObj['AuthorId'] = mapper(cell, users) }

            if (i === 9.0) { xObj['OwnerId'] = mapper(cell, users) }

            if (i === 10.0) { xObj['ComponentId'] = mapper(cell, reqs.dropdowns['Components']) }

            //if empty add null otherwise add the cell to the object under the proper key relative to its location on the template
            //Offset by 2 for proj name and indent level
            //this only handles values for a couple of cases and could be refactored out.
            if (cell === "") {
                xObj[reqs.JSON_headings[i]] = null;
            } else {
                xObj[reqs.JSON_headings[i]] = cell;
            }

        } //end standard cell loop

        //if not empty add object
        //entry MUST have a name
        if (xObj.Name) {
            xObj['ProjectName'] = data.templateData.currentProjectName;

            xObjArr.push(xObj);
           
        }

        xObjArr = parentChildSetter(xObjArr);
    } //end object creator loop

    // set up to individually add each requirement to spirateam
    //error flag, set to true on error
    var isError = null;
    //error log, holds the HTTP error response values
    var errorLog = [];

    //loop through objects to send
    var len = xObjArr.length;
    for (var i = 0; i < len; i++) {
        //stringify
        var JSON_body = JSON.stringify(xObjArr[i]);

        //send JSON, project number, current user data, and indent position to export function
        var response = requirementExportCall(JSON_body, data.templateData.currentProjectNumber, data.userData.currentUser, xObjArr[i].positionNumber);

        //parse response
        if (response.getResponseCode() === 200) {
            //get body information
            response = JSON.parse(response.getContentText())
            responses.push(response.RequirementId)
                //set returned ID to id field
            xObjArr[i].idField.setValue(response.RequirementId)

            //modal that displays the status of each artifact sent
            htmlOutputSuccess = HtmlService.createHtmlOutput('<p>' + (i + 1) + ' of ' + (len) + ' sent!</p>').setWidth(250).setHeight(75);
            SpreadsheetApp.getUi().showModalDialog(htmlOutputSuccess, 'Progress');
        } else {
            //push errors into error log
            errorLog.push(response.getContentText());
            isError = true;
            //set returned ID
            //removed by request can be added back if wanted in future versions
            //xObjArr[i].idField.setValue('Error')

            //Sets error HTML modal
            htmlOutput = HtmlService.createHtmlOutput('<p>Error for ' + (i + 1) + ' of ' + (len) + '</p>').setWidth(250).setHeight(75);
            SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Progress');
        }
    }
    //return the error flag and array with error text responses
    return [isError, errorLog];
}

//Post API call
//takes the stringifyed object, project number, current user, and the position number
function requirementExportCall(body, projNum, currentUser, posNum) {
    //encryption
    var decoded = Utilities.base64Decode(currentUser.api_key);
    var APIKEY = Utilities.newBlob(decoded).getDataAsString();

    //unique url for requirement POST
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + projNum + '/requirements/indent/' + posNum + '?username=';
    //build URL for fetch
    var URL = currentUser.url + fetcherURL + currentUser.userName + APIKEY;
    //POST headers
    var init = {
        'method': 'post',
        'contentType': 'application/json',
        'muteHttpExceptions': true,
        'payload': body
    };

    //calls and returns google fetch function
    return UrlFetchApp.fetch(URL, init);
}


//map cell data to their corresponding IDs for export to spirateam
function mapper(item, list) {
    //set return value to 1 on err
    var val = 1;
    //loop through model for variable being mapped
    for (var i = 0; i < list.length; i++) {
        //cell value matches model value assign id number
        if (item == list[i][1]) { val = list[i][0] }
    }
    return val;
}

//gets full model data and custom properties cell range
function customHeaderRowBuilder(data, rowRange) {
    //shorten variables
    var customs = data.templateData.requirements.customFields;
    var users = data.userData.projUserWNum;
    //length of custom data to optimize perf
    var len = customs.length;
    //custom props array of objects to be returned
    var customProps = [];
    //loop through cells based on custom data fields
    for (var i = 0; i < len; i++) {
        //assign custom property to variable
        var customData = customs[i];
        //get cell data
        var cell = rowRange.offset(0, i).getValue()
            //check if the cell is empty
        if (cell !== "") {
            //call custom content function and push data into array from export
            customProps.push(customFiller(cell, customData, users))
        }
    }
    //custom properties array ready for API export
    return customProps
}

//gets specific cell and custom property data for that column
function customFiller(cell, data, users) {
    //all custom values need a property number
    //set it and add to object for return
    var propNum = data.PropertyNumber;
    var prop = { PropertyNumber: propNum }

    //check data type of custom fields and assign values if condition is met
    if (data.CustomPropertyTypeName == 'Text') {
        prop['StringValue'] = cell;
    }

    if (data.CustomPropertyTypeName == 'Integer') {
        //removes floating points
        cell = parseInt(cell);
        prop['IntegerValue'] = cell;
    }

    if (data.CustomPropertyTypeName == 'Decimal') {
        prop['DecimalValue'] = cell;
    }

    if (data.CustomPropertyTypeName == 'Boolean') {
        //google cells wouldn't validate 'true' or 'false', I assume they're reserved keywords.
        //Used yes and no instead and here they are converted to true and false;
        cell == "Yes" ? prop['BooleanValue'] = true : prop['BooleanValue'] = false;
    }

    if (data.CustomPropertyTypeName == 'List') {
        var len = data.CustomList.Values.length;
        //loop through custom list and match name to cell value
        for (var i = 0; i < len; i++) {
            if (cell == data.CustomList.Values[i].Name) {
                //assign list value number to integer
                prop['IntegerValue'] = data.CustomList.Values[i].CustomPropertyValueId
            }
        }
    }

    if (data.CustomPropertyTypeName == 'Date') {
        //parse date into milliseconds
        cell = Date.parse(cell);
        //concat values accepted by spira and assign to correct prop
        prop['DateTimeValue'] = "\/Date(" + cell + ")\/";
    }


    if (data.CustomPropertyTypeName == 'MultiList') {
        //TODO add some sort of multiList functionality
        //currently 4/2017 Google app script does not support multi select on google sheets

        //single item exported in an array
        var listArray = [];
        var len = data.CustomList.Values.length;
        //loop through custom list and match name to cell value
        for (var i = 0; i < len; i++) {
            if (cell == data.CustomList.Values[i].Name) {
                //assign list value number to integer
                listArray.push(data.CustomList.Values[i].CustomPropertyValueId)
                prop['IntegerListValue'] = listArray;
            }
        }
    }

    if (data.CustomPropertyTypeName == 'User') {
        //loop through users list and assign the id to the property value
        var len = users.length
        for (var i = 0; i < len; i++) {
            if (cell == users[i][1]) {
                prop['IntegerValue'] = users[i][0];
            }
        }
    }

    //return prop object with id and correct value
    return prop;
}

//This function counts the number of '>'s and returns the value
function indenter(cell) {
    var indentCount = 0;
    //check for cell value and indent character '>'
    if (cell && cell[0] === '>') {
        //increment indent counter while there are '>'s present
        while (cell[0] === '>') {
            //get entry length for slice
            var len = cell.length;
            //slice the first character off of the entry
            cell = cell.slice(1, len);
            indentCount++;
        }
    }
    return indentCount
}

function parentChildSetter(arr) {
    //takes the entire array of objects to be sent
    var len = arr.length;
    //this acts as the indent reset
    //when this is 0 it means that the object has a '0' indent level, meaning it should be sitting at the root level (far left)
    var location = 0;

    //loop through the export array
    for (var i = 0; i < len; i++) {
        //if the object has an indent level and the level IS NOT the same as the previous object
        if (arr[i].indentCount > 0 && arr[i].indentCount !== location) {
            //change the position number
            //this can be negative or positive
            arr[i].positionNumber = arr[i].indentCount - location;

            //set the current location for the next object in line
            location = arr[i].indentCount;
        }

        //if the object DOES NOT have an indent level. For example the object is sitting at the root or there was an entry error.
        if (arr[i].indentCount == 0) {
            //this is a hack to push the object all the way to the root position. Currently the API does not support a call to place an artifact at a certain location.
            arr[i].positionNumber = -10;
            //reset location variable
            location = 0;
        }
    }
    //return indented array
    return arr;
}










/*
 * ===================
 * RETRIEVE FROM SPIRA
 * ===================
 * 
 * Chris: This function was initially started and abandoned as the project goals were realigned.
 * Chris: Everything listed below can be safely removed and or modified without effecting the codebase.
 *
 */

//import function, basic for now
//stretch goal is to have this as a useful import
function importer(currentUser, data){
  //get spreadsheet and active first sheet
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadSheet.getSheets()[0];

  // needed for eventual actual importer
  // var paramsCount = '/services/v5_0/RestService.svc/projects/1/requirements/count?username=';
  // var count = getFetch(currentUser, paramsCount )

  //call defined fetch function
  //current params has count set to 35, this can be set/changed programmatically with the count call listed above (stretch goal)
  var params = '/services/v5_0/RestService.svc/projects/1/requirements?starting_row=1&number_of_rows=35&username=';
  var data = fetcher(currentUser, params)

  //get first row range
  var range = sheet.getRange(data.templateData.requirements.editableRange);

  //loop through cells in range
  for(var i = 0; i < data.length; i++){
    var spreadSheet_i = i + 1
    range.getCell(spreadSheet_i, 1).setValue(data[i].RequirementId);
    range.getCell(spreadSheet_i, 2).setValue(data[i].Name);
    range.getCell(spreadSheet_i, 3).setValue(data[i].Description);
    range.getCell(spreadSheet_i, 4).setValue(data[i].ReleaseVersionNumber);
    range.getCell(spreadSheet_i, 5).setValue(data[i].RequirementTypeName);
    range.getCell(spreadSheet_i, 6).setValue(data[i].ImportanceName);
    range.getCell(spreadSheet_i, 7).setValue(data[i].StatusName);
    range.getCell(spreadSheet_i, 8).setValue(data[i].EstimatePoints);
    range.getCell(spreadSheet_i, 9).setValue(data[i].AuthorName);
    range.getCell(spreadSheet_i, 10).setValue(data[i].OwnerName);
    range.getCell(spreadSheet_i, 11).setValue(data[i].ComponentId);

    //moves the range down one row
    range = range.offset(1, 0, data.templateData.requirements.cellRangeLength);
 }
}