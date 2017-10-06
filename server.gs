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
 * ====================
 * DATA "GET" FUNCTIONS
 * ====================
 * 
 * functions used to retrieve data from Spira - things like projects and users, not specific records
 * 
 */

// General fetch function, using Google's built in fetch api
// @param: currentUser - user object storing login data from client
// @param: fetcherUrl - url string passed in to connect with Spira
function fetcher(currentUser, fetcherURL) { 
    //google base 64 encoded string utils
    var decoded = Utilities.base64Decode(currentUser.api_key);
    var APIKEY = Utilities.newBlob(decoded).getDataAsString();

    //build URL from args
    var fullUrl = currentUser.url + fetcherURL + "username=" + currentUser.userName + APIKEY;
    //set MIME type
    var params = { 'content-type': 'application/json' };
    
    //call Google fetch function
    var response = UrlFetchApp.fetch(fullUrl, params);
    
    //returns parsed JSON
    //unparsed response contains error codes if needed
    return JSON.parse(response);
}



// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
// @param: currentUser - object with details about the current user
function getProjects(currentUser) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects?';
    return fetcher(currentUser, fetcherURL);
}



// Gets components for selected project.
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getComponents(currentUser, projectId) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + projectId + '/components?active_only=true&include_deleted=false&';
    return fetcher(currentUser, fetcherURL);
}



// Gets custom fields for selected project and artifact
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: artifactName - string name of the current artifact
function getCustoms(currentUser, projectId, artifactName) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + projectId + '/custom-properties/' + artifactName + '?';
    return fetcher(currentUser, fetcherURL);
}



// Gets releases for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getReleases(currentUser, projectId) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + projectId + '/releases?';
    return fetcher(currentUser, fetcherURL);
}



// Gets users for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getUsers(currentUser, projectId) {
    var fetcherURL = '/services/v5_0/RestService.svc/projects/' + projectId + '/users?';
    return fetcher(currentUser, fetcherURL);
}









/*
 *
 * =======================
 * CREATE "POST" FUNCTIONS
 * =======================
 * 
 * functions to create new records in Spira - eg add new requirements
 * 
 */

// General fetch function, using Google's built in fetch api
// @param: body - json object
// @param: currentUser - user object storing login data from client
// @param: postUrl - url string passed in to connect with Spira
function poster(body, currentUser, postUrl) {
    //encryption
    var decoded = Utilities.base64Decode(currentUser.api_key);
    var APIKEY = Utilities.newBlob(decoded).getDataAsString();

    //build URL from args
    var fullUrl = currentUser.url + postUrl + "username=" + currentUser.userName + APIKEY;

    //POST headers
    var params = {
        'method': 'post',
        'contentType': 'application/json',
        'muteHttpExceptions': true,
        'payload': body
    };

    //call Google fetch function
    var response = UrlFetchApp.fetch(fullUrl, params);
    
    //returns parsed JSON
    //unparsed response contains error codes if needed
    return response;
    //return JSON.parse(response);
}



// post new requirement
// @param: body - stringified object of all relevant fields
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: indentPosition - int used for setting the relative indenting position
function postRequirementToSpira(body, currentUser, projectId, indentPosition) {
    //unique url for requirement POST
    var postUrl = '/services/v5_0/RestService.svc/projects/' + projectId + '/requirements/indent/' + indentPosition + "?";
    
    poster(body, currentUser, postUrl);
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
// @param: string - message to be displayed
function warn(string) {
    var ui = SpreadsheetApp.getUi();
    //alert popup with yes and no button
    var response = ui.alert(string, ui.ButtonSet.YES_NO);

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



// Google alert popup with OK button
// @param: dialog - message to show
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

// function that manages template creation - creating the header row, formatting cells, setting validation
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldType - list of fieldType enums from client params object
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
                    list.push(model.projectUsers[i].name);
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

// function that manages exporting data from the sheet - creating an array of objects based on entered data, then sending to Spira
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldType - list of fieldType enums from client params object
function exporter(model, fieldType) {
    // get the active spreadsheet and first sheet
    // TODO rework this to be the active sheet - not the first one
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
        sheet = spreadSheet.getSheets()[0],
        fields = model.fields,
        artifact = model.currentArtifact,
        artifactIsHierarchical = artifact.hierarchical,
        artifactHasFolders = artifact.hasFolders,

        sheetData = sheet.getRange(2,1, sheet.getLastRow() - 1, fields.length).getValues(),
        entriesForExport = new Array;

    
    var lastIndentPosition = "";
    for (var row = 0; row < sheetData.length; row++) {
        // stop at the first row that is fully blank
        if (!sheetData[row].join() || !rowHasRequiredFields(sheetData[row], fields)) {
            break;
        } else {
            var entry = createEntryFromRow( sheetData[row], model, fieldType, artifactIsHierarchical, lastIndentPosition );
            entriesForExport.push(entry);
            
            // update the last indent position before going to the next entry to make sure relative indent is set correctly
            if (artifactIsHierarchical) {
                lastIndentPosition = ( entry.indentPosition < 0 ) ? 0 : entry.indentPosition;
            }
        }
    }
  
    return entriesForExport;

    // Create and show a window to tell the user what is going on
    var exportMessageToUser = HtmlService.createHtmlOutput('<p>Preparing your data for export!</p>').setWidth(250).setHeight(75);
    SpreadsheetApp.getUi().showModalDialog(exportMessageToUser, 'Progress');



    // set up to individually add each requirement to spirateam
    //error flag, set to true on error
    var isError = null;
    //error log, holds the HTTP error response values
    var errorLog = [];

    //loop through objects to send
    var len = entriesForExport.length;
    for (var i = 0; i < len; i++) {
        //stringify
        var JSON_body = JSON.stringify(entriesForExport[i]);

        //send JSON object of new requirement, current user data, project number, and indent position to export function
        var response = postRequirementToSpira(
            JSON_body, 
            model.user, 
            model.currentProject.id, 
            entriesForExport[i].indentPosition
        );
      
       return response;

        //parse response
        if (response.getResponseCode() === 200) {
            //get body information
            response = JSON.parse(response.getContentText())
            responses.push(response.RequirementId)
                //set returned ID to id field
            entriesForExport[i].idField.setValue(response.RequirementId)

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






// check to see if a row of data has entries for all required fields
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
function rowHasRequiredFields(row, fields) {
    var result = true;
    for (var column = 0; column < row.length; column++) {
        if (fields[column].required && !row[column]) {
            result = false;
        }
    }
    return result;
}



// function creates a correctly formatted artifact object ready to send to Spira
// it works through each field type to validate and parse the values so object is in correct form
// any field that does not pass validation receives a null value
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: model - full model with info about fields, dropdowns, users, etc
// @param: fieldType - object of all field types with enums
// @param: artifactIsHierarchical - bool to tell function if this artifact has hierarchy (eg RQ and RL)
// @param: lastIndentPosition - global int used for calculating relative indents for hierarchical artifacts
function createEntryFromRow(row, model, fieldType, artifactIsHierarchical, lastIndentPosition) {
    //create empty 'entry' object - include custom properties array here to avoid it being undefined later if needed
    var entry = {
            "CustomProperties": new Array
        },
        fields = model.fields;

    //we need to turn an array of values in the row into a validated object 
    for (var index = 0; index < row.length; index++) {
        var value = null,
            customType = "";

        // double check data validation, convert dropdowns to required int values
        // sets both the value, and custom types - so that custom fields are handled correctly
        switch (fields[index].type) {
            
            // ID fields: restricted to numbers and blank on push, otherwise put
            case fieldType.id:

                customType = "IntegerValue";
                break;
            
            // INT and NUM fields are both treated by Sheets as numbers
            case fieldType.int:
            case fieldType.num:
                // only set the value if a number has been returned
                if (!isNaN(row[index])) {
                    value = row[index];
                    customType = "IntegerValue";
                };
                break;

            // BOOL as Sheets has no bool validation, a yes/no dropdown is used
            case fieldType.bool:
                // 'True' and 'False' don't work as dropdown choices, so have to convert back
                if (row[index] == "Yes") {
                    value = true;
                    customType = "BooleanValue";
                } else if (row[index] == "No") {
                    value = false;
                    customType = "BooleanValue";
                };
                break;

            // DATES - parse the data and add prefix/suffix for WCF
            case fieldType.date:
                if (row[index]) {
                    value = "\/Date(" + Date.parse(row[index]) + ")\/";
                    customType = "DateTimeValue";
                }
                break;

            // DROPDOWNS - get id from relevant name, if one is present
            case fieldType.drop:
                var idFromName = getIdFromName(row[index], fields[index].values);
                if (idFromName) {
                    value = idFromName;
                    customType = "IntegerValue";
                }
                break;

            // MULTIDROPDOWNS - get id from relevant name, if one is present, set customtype to list value
            case fieldType.multi:
                var idFromName = getIdFromName(row[index], fields[index].values);
                if (idFromName) {
                    value = idFromName;
                    customType = "IntegerListValue";
                }
                break;

            // USER fields - get id from relevant name, if one is present
            case fieldType.user:
                var idFromName = getIdFromName(row[index], model.projectUsers);
                if (idFromName) {
                    value = idFromName;
                    customType = "IntegerValue";
                }
                break;

            // COMPONENT fields - get id from relevant name, if one is present
            case fieldType.component:
                var idFromName = getIdFromName(row[index], model.projectComponents);
                if (idFromName) {
                    value = idFromName;
                    customType = "IntegerValue";
                }
                break;
              
            // RELEASE fields - get id from relevant name, if one is present
            case fieldType.release:
                var idFromName = getIdFromName(row[index], model.projectReleases);
                if (idFromName) {
                    value = idFromName;
                    customType = "IntegerValue";
                }
                break;
            
            // All other types
            default:
                // just assign the value to the cell - used for text
                value = row[index];
                customType = "StringValue";
                break;
        }


        // handle hierarchy fields - if required: checks artifact type is hierarchical and if this field sets hierarchy
        if (artifactIsHierarchical && fields[index].setsHierarchy) {
            // first get the number of indent characters
            var indentCount = countAndRemoveIndentCharacaters(value, indentCharacter);
            var indentPosition = setRelativePosition(indentCount, lastIndentPosition);
            
            // make sure to slice off the indent characters from the front
            // TODO should also trim white space at start
            value = value.slice(indentCount, value.length);

            // set the indent position for this row
            entry.indentPosition = indentPosition;
        }

        // check whether field is marked as a custom field and as the required property number 
        if (fields[index].isCustom && fields[index].propertyNumber) {

            // if field has data create the object
            if (value) {
                var customObject = {};
                customObject.PropertyNumber = fields[index].propertyNumber;
                customObject[customType] = value;

                entry.CustomProperties.push(customObject);
            }
          
        // add standard fields in standard way - only add if field contains data
        } else if (value) {
            entry[fields[index].field] = value;
        }
    }

    return entry;
}



// find the corresponding ID for a string value - eg from a dropdown
// dropdowns can only contain one item per row so we have to now get the IDs for sending to Spira
// @param: string - the string of the name value specified
// @param: list - the array of items with keys for id and name values
function getIdFromName(string, list) {
    for (var i = 0; i < list.length; i++) {
        if (list[i].name == string) { 
            return list[i].id;
        }
    }
    // return 0 if there's no match
    return 0;
}



// returns the count of the number of indent characters and returns the value
// @param: field - a single field string - one already designated as containing hierarchy info
// @param: indentCharacter - the character used to denote an indent - e.g. ">"
function countIndentCharacaters(field, indentCharacter) {
    var indentCount = 0;
    //check for field value and indent character
    if (field && field[0] === indentCharacter) {
        //increment indent counter while there are '>'s present
        while (field[0] === indentCharacter) {
            //get entry length for slice
            var len = field.length;
            //slice the first character off of the entry
            field = field.slice(1, len);
            indentCount++;
        }
    }
    return indentCount;
}



// returns the correct relative indent position - based on the previous relative indent and other logic
// @param: indentCount - int of the number of indent characters set by user
// @param: lastIndentPosition - int of the actual indent position used for the preceding entry/row
function setRelativePosition(indentCount, lastIndentPosition) {
    if (indentCount !== lastIndentPosition) {
        // set the indent positions relative to the previous position (can be negative or positive)
        return indentCount - lastIndentPosition;

    } else if (indentCount === lastIndentPosition) {
        // set the indent positions to be the same as the last
        return lastIndentPosition;

    } else if (indentCount === 0) {
        // this is a hack to push item (hopefully) all the way to the root position. 
        // Currently the API does not support a call to place an artifact at a certain location.
        return -10;

    } else {
        // otherwise just set it to zero
        return 0;
    }
}





//DONT THINK NEED THIS ANYMORE BUT HERE IN CASE MY CODE IS WRONG
/*
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
}*/










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