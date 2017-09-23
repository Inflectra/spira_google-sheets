/*
 * ======================
 * INITIAL LOAD FUNCTIONS
 * ======================
 * 
 * These functions are needed for initialization  
 * All Google App Script (GAS) files are bundled by the engine 
 * at start up so any non-scoped variables declared will be available globally.
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
    .setTitle('SpiraTeam by Inflectra');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}



// This function is part of the google template engine and allows for modularization of code
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}









/*
 *
 * ========================
 * TEMPLATE PANEL FUNCTIONS
 * ========================
 * 
 */

// copy the first sheet into a new sheet in the same spreadsheet
function save() {
    // pop up telling the user that their data will be saved
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('This will save the current sheet in a new sheet on this spreadsheet. Continue?', ui.ButtonSet.YES_NO);

    // returns with user choice
    if (response == ui.Button.YES) {
        // get first sheet of  active spreadsheet
        var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = spreadSheet.getSheets()[0];

        // get entire open spreadsheet id
        var id = spreadSheet.getId();

        // set current spreadsheet file as destination
        var destination = SpreadsheetApp.openById(id);

        // copy sheet to current spreadsheet in new sheet
        sheet.copyTo(destination);
      
        // returns true to queue success popup
        return true;
    } else {
        // returns false to ignore success popup     
        return false;
    }
}



//clears first sheet in spreadsheet
function clearAll() {
    // get first active spreadsheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getSheets()[0];

    // clear all formatting and content
    sheet.clear();

    // clears data validations from the entire sheet
    var range = SpreadsheetApp.getActive().getRange('A:AZ');
    range.clearDataValidations();

    // Reset sheet name
    sheet.setName('Sheet');
}









/*
 *
 * =====================
 * FETCH "GET" FUNCTIONS
 * =====================
 * 
 */

// General fetch function, using Google's built in fetch api
// @param: currentUser = user object storing login data from client
// @param: fetcherUrl = url string passed in to connect with Spira
function fetcher(currentUser, fetcherURL) { 
    //google base 64 encoded string utils
    var decoded = Utilities.base64Decode(currentUser.api_key);
    var APIKEY = Utilities.newBlob(decoded).getDataAsString();

    //build URL from args
    var URL = currentUser.url + fetcherURL + currentUser.userName + APIKEY;
    //set MIME type
    var params = { 'content-type': 'application/json' };
    //call Google fetch function
    var response = UrlFetchApp.fetch(URL, params);
    
    //returns parsed JSON
    //unparsed response contains error codes if needed
    return JSON.parse(response);
}



// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
function getProjects(currentUser) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects?username=';
    return fetcher(currentUser, fetcherURL);
}



// Gets components for selected project.
function getComponents(currentUser, proj) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/components?active_only=true&include_deleted=false&username=';
    return fetcher(currentUser, fetcherURL);
}



// Gets custom fields for selected project and artifact
function getCustoms(currentUser, proj, artifact) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/custom-properties/' + artifact + '?username=';
    return fetcher(currentUser, fetcherURL);
}



// Gets releases for selected project.
function getReleases(currentUser, proj) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/releases?username=';
    return fetcher(currentUser, fetcherURL);
}



// Gets users for selected project
function getUsers(currentUser, proj) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + proj + '/users?username=';
    return fetcher(currentUser, fetcherURL);
}









/*
 *
 * ==============
 * ERROR MESSAGES
 * ==============
 * 
 */

// Error notification function
// Assigns string value and routes error call from client.js.html
// @param: type - string identifying the message to be displayed
function error(type) {
    if (type == 'impExp') {
        okWarn('There was an input error. Please check that your entries are correct.');
    } else if (type == 'unknown') {
        okWarn('Unkown error. Please try again later or contact your system administrator');
    } else {
        okWarn('Network error. Please check your username, url, and password. If correct make sure you have the correct permissions.');
    }
}



// Pop-up notification function
// @param: string - message to be displayed
function success(string) {
    // Show a 2-second popup with the title "Status" and a message passed in as an argument.
    SpreadsheetApp.getActiveSpreadsheet().toast(string, 'Success', 2);
}



// Alert pop up for data clear warning
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



// Alert pop up for export success
// @param: err - boolean sent from the export function
function exportSuccess(err) {
    if (err) {
        okWarn('Operation complete, some errors occurred. Clear sheet to export more artifacts.');
    } else {
        okWarn('Operation complete. Clear sheet to export more artifacts.');
    }
}

// Alert pop up for no template present
function noTemplate() {
    okWarn('Please load a template to continue.');
}

// Google alert popup with Ok button
function okWarn(dialog) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(dialog, ui.ButtonSet.OK);
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

// function for template creation
function templateLoader(model, fieldType) {
    // clear spreadsheet depending on user input
    clearAll();

    // select open file and select first sheet
    // TODO rework this to be the active sheet - not the first one
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
        sheet = spreadSheet.getSheets()[0],
        fields = model.fields;
    
    // set sheet (tab) name to model name
    sheet.setName(model.currentProject.name + ' - ' + model.currentArtifact.name);
    
    // heading row - sets names and formatting
    headerSetter(sheet, fields, model.colors);

    // set validation rules on the columns
    contentValidationSetter(sheet, model, fieldType);

    // set any extra formatting options
    contentFormattingSetter(sheet, model);

    /*
    //loop through model size data and set columns to correct width
    for (var i = 0; i < artifactData.sizes.length; i++) {
        sheet.setColumnWidth(artifactData.sizes[i][0], artifactData.sizes[i][1]);
    }
    */
}



// Sets headings for fields
// creates an array of the field names so that changes can be batched to the relevant range in one go for performance reasons
// @param: sheet - the sheet object
// @param: fields - full field data
// @param: colors - global colors used for formatting
function headerSetter (sheet, fields, colors) {
    
    var headerNames = [],
        fieldsLength = fields.length;

    for (var i = 0; i < fieldsLength; i++) {
        headerNames.push(fields[i].name);
    }

    sheet.getRange(1, 1, 1, fieldsLength)
        .setWrap(true)
        .setBackground(colors.bgHeader)
        .setFontColor(colors.header)
        // the headerNames array needs to be in an array as setValues expects a 2D for managing 2D ranges
        .setValues([headerNames])
        .protect().setDescription("header row").setWarningOnly(true);
}



// Sets validation on a per column basis, based on the field type passed in by the model
// a switch statement checks for any type requiring validation and carries out necessary action
// @param: sheet - the sheet object
// @param: model - full data to acccess global params as well as all fields
// @param: fieldType - enums for field types
function contentValidationSetter (sheet, model, fieldType) {
    for (var index = 0; index < model.fields.length; index++) {
        var columnNumber = index + 1;
        
        switch (model.fields[index].type) {
            
            // ID fields: restricted to numbers and protected
            case fieldType.id:
                setNumberValidation(sheet, columnNumber, model.rowsToFormat, false);
                protectColumn(
                    sheet, 
                    columnNumber, 
                    model.rowsToFormat, 
                    model.colors.bgReadOnly, 
                    "ID field",
                    false
                    );
                break;
            
            // INT and NUM fields are both treated by Sheets as numbers
            case fieldType.int:
            case fieldType.num:
                setNumberValidation(sheet, columnNumber, model.rowsToFormat, false);
                break;

            // BOOL as Sheets has no bool validation, a yes/no dropdown is used
            case fieldType.bool:
                // 'True' and 'False' don't work as dropdown choices
                var list = ["Yes", "No"];
                setDropdownValidation(sheet, columnNumber, model.rowsToFormat, list, false);
                break;

            case fieldType.date:
                setDateValidation(sheet, columnNumber, model.rowsToFormat, false);
                break;

            // DROPDOWNS and MULTIDROPDOWNS are both treated as simple dropdowns (Sheets does not have multi selects)
            case fieldType.drop:
            case fieldType.multi:
                var list = [];
                var fieldList = model.fields[index].values;
                for (var i = 0; i < fieldList.length; i++) {
                    list.push(fieldList[i].name);
                }
                setDropdownValidation(sheet, columnNumber, model.rowsToFormat, list, false);
                break;

            // USER fields are dropdowns with the values coming from a project wide set list
            case fieldType.user:
                var list = [];
                for (var i = 0; i < model.projectUsers.length; i++) {
                    list.push(model.projectUsers[i].fullName);
                }
                setDropdownValidation(sheet, columnNumber, model.rowsToFormat, list, false);
                break;

            // COMPONENT fields are dropdowns with the values coming from a project wide set list
            case fieldType.component:
                var list = [];
                for (var i = 0; i < model.projectComponents.length; i++) {
                    list.push(model.projectComponents[i].name);
                }
                setDropdownValidation(sheet, columnNumber, model.rowsToFormat, list, false);
                break;
              
            // RELEASE fields are dropdowns with the values coming from a project wide set list
            case fieldType.release:
                var list = [];
                for (var i = 0; i < model.projectReleases.length; i++) {
                    list.push(model.projectReleases[i].name);
                }
                setDropdownValidation(sheet, columnNumber, model.rowsToFormat, list, false);
                break;
            
            // All other types
            default:
                //do nothing
                break;
        }
    }
}



// create dropdown validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: list - array of values to show in a dropdown and use for validation
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setDropdownValidation (sheet, columnNumber, rowLength, list, allowInvalid) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    // requireValueInList - params are the array to use, and whether to create a dropdown list
    var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(list, true)
        .setAllowInvalid(allowInvalid)
        .build();
    range.setDataValidation(rule);
}



// create date validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setDateValidation (sheet, columnNumber, rowLength, allowInvalid) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    var rule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .setHelpText('Must be a valid date')
        .build();
    range.setDataValidation(rule);
}



// create number validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setNumberValidation (sheet, columnNumber, rowLength, allowInvalid) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);

    // create the validation rule
    //must be a valid number greater than -1 (also excludes 1.1.0 style numbers)
    var rule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(-1)
        .setAllowInvalid(allowInvalid)
        .setHelpText('Must be a positive number')
        .build();
    range.setDataValidation(rule);
}


// format columns based on a potential rang of factors - eg hide unsupported columns
// @param: sheet - the sheet object
// @param: model - full model data set
function contentFormattingSetter (sheet, model) {
    for (var i = 0; i < model.fields.length; i++) {
        var columnNumber = i + 1;
        
        // hide unsupported fields
        if (model.fields[i].unsupported) {
            protectColumn(
              sheet, 
              columnNumber, 
              model.rowsToFormat, 
              model.colors.bgReadOnly, 
              model.fields[i].name + "unsupported",
              true
              );
        }
    }
}



// protects specific column. Edits still allowed - current user not excluded from edit list, but could in future
// @param: sheet - the sheet object
// @param: columnNumber - int of column to hide
// @param: rowLength - int of default number of rows to apply any formattting to
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
// @param: name - string description for the protected range
// @param: hide - optional bool to hide column completely
function protectColumn (sheet, columnNumber, rowLength, bgColor, name, hide) {
    // create range
    var range = sheet.getRange(2, columnNumber, rowLength);
    range.setBackground(bgColor)
        .protect()
        .setDescription(name)
        .setWarningOnly(true);

  if(hide) {
    sheet.hideColumns(columnNumber);
  }
    
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

function exporter(model) {
    // get the active spreadsheet and first sheet
    // TODO rework this to be the active sheet - not the first one
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
        sheet = spreadSheet.getSheets()[0],
        fields = model.fields;

    // full area on the sheet where data may be
    var range = sheet.getRange(1, 1, model.rowsToFormat, fields.length);

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

    // Create and show a window to tell the user what is going on
    var exportMessageToUser = HtmlService.createHtmlOutput('<p>Preparing your data for export!</p>').setWidth(250).setHeight(75);
    SpreadsheetApp.getUi().showModalDialog(exportMessageToUser, 'Progress');

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

            //pass values to getIdFromName function
            //getIdFromName iterates and assigns the values number based on the list order
            if (i === 3.0) { xObj['ReleaseId'] = getIdFromName(cell, reqs.dropdowns['Version Number']) }

            if (i === 4.0) { cell = getIdFromName(cell, reqs.dropdowns['Type']) }

            if (i === 5.0) { xObj['ImportanceId'] = getIdFromName(cell, reqs.dropdowns['Importance']) }

            if (i === 6.0) { xObj['StatusId'] = getIdFromName(cell, reqs.dropdowns['Status']) }

            if (i === 8.0) { xObj['AuthorId'] = getIdFromName(cell, users) }

            if (i === 9.0) { xObj['OwnerId'] = getIdFromName(cell, users) }

            if (i === 10.0) { xObj['ComponentId'] = getIdFromName(cell, reqs.dropdowns['Components']) }

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


// find the corresponding ID for a string value - eg from a dropdown
// dropdowns can only contain one item per row so we have to now get the IDs for sending to Spira
// @param: name - the string of the name value specified
// @param: list - the array of items with keys for id and name values
function getIdFromName(name, list) {
    for (var i = 0; i < list.length; i++) {
        if (item == list[i].name) { 
            return list[i].id;
        }
    }
    // return 0 if there's no match
    return 0;
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
/*
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
*/