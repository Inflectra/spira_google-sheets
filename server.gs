// globals
var API_BASE = '/services/v5_0/RestService.svc/projects/',
    API_BASE_NO_SLASH = '/services/v5_0/RestService.svc/projects',
    ART_ENUMS = {
        requirements: 1,
        testCases: 2,
        incidents: 3,
        releases: 4,
        testRuns: 5,
        tasks: 6,
        testSteps: 7,
        testSets: 8
    },
    FIELD_MANAGEMENT_ENUMS = {
        all: 1,
        standard: 2,
        subType: 3
    },
    STATUS_ENUM = {
        allSuccess: 1,
        allError: 2,
        someError: 3
    },
    INLINE_STYLING = "style='font-family: sans-serif'";

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
    var lastColumn = sheet.getMaxColumns(),
        lastRow = sheet.getMaxRows();
    sheet.clear(); //.showColumns(1,  lastColumn)

    // clears data validations and notes from the entire sheet
    var range = sheet.getRange(1, 1, lastRow, lastColumn);
    range.clearDataValidations().clearNote();

    // remove any protections on the sheet
    var protections = spreadSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
       var protection = protections[i];
       if (protection.canEdit()) {
           protection.remove();
       }
   }

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
    var fetcherURL = API_BASE_NO_SLASH + '?';
    return fetcher(currentUser, fetcherURL);
}



// Gets components for selected project.
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getComponents(currentUser, projectId) {
    var fetcherURL = API_BASE + projectId + '/components?active_only=true&include_deleted=false&';
    return fetcher(currentUser, fetcherURL);
}



// Gets custom fields for selected project and artifact
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: artifactName - string name of the current artifact
function getCustoms(currentUser, projectId, artifactName) {
    var fetcherURL = API_BASE + projectId + '/custom-properties/' + artifactName + '?';
    return fetcher(currentUser, fetcherURL);
}



// Gets releases for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getReleases(currentUser, projectId) {
    var fetcherURL = API_BASE + projectId + '/releases?';
    return fetcher(currentUser, fetcherURL);
}



// Gets users for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getUsers(currentUser, projectId) {
    var fetcherURL = API_BASE + projectId + '/users?';
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



// effectively a switch to manage which artifact we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactId - int of the current artifact
// @param: parentId - optional int of the relevant parent to attach the artifact too
function postArtifactToSpira(entry, user, projectId, artifactId, parentId) {

    //stringify
    var JSON_body = JSON.stringify(entry),
        response = "",
        postUrl = "";

    //send JSON object of new item to artifact specific export function
    switch (artifactId) {

        // REQUIREMENTS
        case ART_ENUMS.requirements:
            postUrl = API_BASE + projectId + '/requirements/indent/' + entry.indentPosition + '?';
            response = poster(JSON_body, user, postUrl);
            break;

        // TEST CASES
        case ART_ENUMS.testCases:
            postUrl = API_BASE + projectId + '/test-cases?';
            response = poster(JSON_body, user, postUrl);
            break;

        // INCIDENTS
        case ART_ENUMS.incidents:
            postUrl = API_BASE + projectId + '/incidents?';
            response = poster(JSON_body, user, postUrl);
            break;

        // RELEASES
        case ART_ENUMS.releases:
            postUrl = API_BASE + projectId + '/releases?';
            response = poster(JSON_body, user, postUrl);
            break;

        // TASKS
        case ART_ENUMS.tasks:
            postUrl = API_BASE + projectId + '/tasks?';
            response = poster(JSON_body, user, postUrl);
            break;

        // TEST STEPS
        case ART_ENUMS.testSteps:
            postUrl = API_BASE + projectId + '/test-cases/' + parentId + '/test-steps?';
            response = poster(JSON_body, user, postUrl);
            break;
    }

    return response;
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
// @param: message - string sent from the export function
function exportSuccess(message) {
    if (message ==  STATUS_ENUM.allSuccess) {
        okWarn("All done! To send more data over, clear the sheet first.");
    } else if (message == STATUS_ENUM.someError) {
        okWarn("Sorry, but there were some problems. Check the notes on the relevant ID field for explanations.");
    } else if (message == STATUS_ENUM.alLError){
        okWarn("We're really sorry, but we couldn't send anything to SpiraTeam - please check notes on the ID fields  for more information.");
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
    // clear spreadsheet depending on user input, and unhide everything
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

}



// Sets headings for fields
// creates an array of the field names so that changes can be batched to the relevant range in one go for performance reasons
// @param: sheet - the sheet object
// @param: fields - full field data
// @param: colors - global colors used for formatting
function headerSetter (sheet, fields, colors) {

    var headerNames = [],
        backgrounds = [],
        fontColors = [],
        fontWeights = [],
        fieldsLength = fields.length;

    for (var i = 0; i < fieldsLength; i++) {
        headerNames.push(fields[i].name);

        // set field text depending on whether is required or not
        var fontColor = (fields[i].required || fields[i].requiredForSubType) ? colors.headerRequired : colors.header;
        var fontWeight = fields[i].required ? 'bold' : 'normal';
        fontColors.push(fontColor);
        fontWeights.push(fontWeight);

        // set background colors based on if it is a subtype only field or not
        var background = fields[i].isSubTypeField ? colors.bgHeaderSubType : colors.bgHeader;
        backgrounds.push(background);
    }

    sheet.getRange(1, 1, 1, fieldsLength)
        .setWrap(true)
        // the arrays need to be in an array as methods expect a 2D array for managing 2D ranges
        .setBackgrounds([backgrounds])
        .setFontColors([fontColors])
        .setFontWeights([fontWeights])
        .setValues([headerNames])
        .protect().setDescription("header row").setWarningOnly(true);
}



// Sets validation on a per column basis, based on the field type passed in by the model
// a switch statement checks for any type requiring validation and carries out necessary action
// @param: sheet - the sheet object
// @param: model - full data to acccess global params as well as all fields
// @param: fieldType - enums for field types
function contentValidationSetter (sheet, model, fieldType) {
    var nonHeaderRows = sheet.getMaxRows() - 1;
    for (var index = 0; index < model.fields.length; index++) {
        var columnNumber = index + 1,
            list = [];

        switch (model.fields[index].type) {

            // ID fields: restricted to numbers and protected
            case fieldType.id:
            case fieldType.subId:
                setNumberValidation(sheet, columnNumber, nonHeaderRows, false);
                protectColumn(
                    sheet,
                    columnNumber,
                    nonHeaderRows,
                    model.colors.bgReadOnly,
                    "ID field",
                    false
                    );
                break;

            // INT and NUM fields are both treated by Sheets as numbers
            case fieldType.int:
            case fieldType.num:
                setNumberValidation(sheet, columnNumber, nonHeaderRows, false);
                break;

            // BOOL as Sheets has no bool validation, a yes/no dropdown is used
            case fieldType.bool:
                // 'True' and 'False' don't work as dropdown choices
                list.push("Yes", "No");
                setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
                break;

            case fieldType.date:
                setDateValidation(sheet, columnNumber, nonHeaderRows, false);
                break;

            // DROPDOWNS and MULTIDROPDOWNS are both treated as simple dropdowns (Sheets does not have multi selects)
            case fieldType.drop:
            case fieldType.multi:
                var fieldList = model.fields[index].values;
                for (var i = 0; i < fieldList.length; i++) {
                    list.push(fieldList[i].name);
                }
                setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
                break;

            // USER fields are dropdowns with the values coming from a project wide set list
            case fieldType.user:
                for (var j = 0; j < model.projectUsers.length; j++) {
                    list.push(model.projectUsers[i].name);
                }
                setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
                break;

            // COMPONENT fields are dropdowns with the values coming from a project wide set list
            case fieldType.component:
                for (var k = 0; k < model.projectComponents.length; k++) {
                    list.push(model.projectComponents[i].name);
                }
                setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
                break;

            // RELEASE fields are dropdowns with the values coming from a project wide set list
            case fieldType.release:
                for (var l = 0; l < model.projectReleases.length; l++) {
                    list.push(model.projectReleases[i].name);
                }
                setDropdownValidation(sheet, columnNumber, nonHeaderRows, list, false);
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
              (sheet.getMaxRows() - 1),
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
    // 1. SETUP FUNCTION LEVEL VARS
    // get the active spreadsheet and first sheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
        sheet = spreadSheet.getSheets()[0],
        fields = model.fields,
        artifact = model.currentArtifact,
        artifactIsHierarchical = artifact.hierarchical,
        artifactHasFolders = artifact.hasFolders,
        lastRow = sheet.getLastRow() - 1 || 10, // hack to make sure we pass in some rows to the sheetRange, otherwise it causes an error

        sheetRange = sheet.getRange(2,1, lastRow, fields.length),
        sheetData = sheetRange.getValues(),
        entriesForExport = [],
        lastIndentPosition = null;



    // 2. CREATE ARRAY OF ENTRIES
    // loop to create artifact objects from each row taken from the spreadsheet
    for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {

        // stop at the first row that is fully blank
        if (sheetData[rowToPrep].join("") === "") {
            break;
        } else {
            // check for required fields (for normal artifacts and those with sub types - eg test cases and steps)
            var rowChecks = {
                    hasSubType: artifact.hasSubType,
                    totalFieldsRequired: countRequiredFieldsByType(fields, false),
                    totalSubTypeFieldsRequired: artifact.hasSubType ? countRequiredFieldsByType(fields, true) : 0,
                    countRequiredFields: rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, false),
                    countSubTypeRequiredFields: artifact.hasSubType ? rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, true) : 0,
                    subTypeIsBlocked: !artifact.hasSubType ? true : rowBlocksSubType(sheetData[rowToPrep], fields)
                },

                // create entry used to populate all relevant data for this row
                entry = {};

            // first check for errors
            var hasProblems = rowHasProblems(rowChecks);
            if (hasProblems) {
                entry.validationMessage = hasProblems;

            // if error free determine what field filtering is required - needed to choose type/subtype fields if subtype is present
            } else {
                var fieldsToFilter = relevantFields(rowChecks);
                entry = createEntryFromRow( sheetData[rowToPrep], model, fieldType, artifactIsHierarchical, lastIndentPosition, fieldsToFilter );

                // FOR SUBTYPE ENTRIES add flag on entry if it is a subtype
                if (fieldsToFilter === FIELD_MANAGEMENT_ENUMS.subType) {
                    entry.isSubType = true;
                }
                // FOR HIERARCHICAL ARTIFACTS update the last indent position before going to the next entry to make sure relative indent is set correctly
                if (artifactIsHierarchical) {
                    lastIndentPosition = ( entry.indentPosition < 0 ) ? 0 : entry.indentPosition;
                }
            }
            entriesForExport.push(entry);
        }
    }


    // 3. GET READY TO SEND DATA TO SPIRA
    // Create and show a window to tell the user what is going on
    if (!entriesForExport.length) {
        var nothingToExportMessage = HtmlService.createHtmlOutput('<p ' + INLINE_STYLING + '>There are no entries to send to SpiraTeam</p>').setWidth(250).setHeight(75);
        SpreadsheetApp.getUi().showModalDialog(nothingToExportMessage, 'Check Sheet');
        return "nothing to send";
    } else {

        var exportMessageToUser = HtmlService.createHtmlOutput('<p ' + INLINE_STYLING + '>Sending to SpiraTeam...</p>').setWidth(150).setHeight(75);
        SpreadsheetApp.getUi().showModalDialog(exportMessageToUser, 'Progress');

        // create required variables for managing responses for sending data to spirateam
        var log = {
                errorCount: 0,
                successCount: 0,
                entriesLength: entriesForExport.length,
                entries: []
            },
            // set var for parent - used to designate eg a test case so it can be sent with the test step post
            parentId = 0;



        // 4. SEND DATA TO SPIRA AND MANAGE RESPONSES
        //loop through objects to send
        for (var i = 0; i < entriesForExport.length; i++) {
          var response = {};

            // skip if there was an error validating the sheet row
            if (entriesForExport[i].validationMessage) {
                response.error = true;
                response.message = entriesForExport[i].validationMessage;
                log.errorCount++;
            }
            // skip if a sub type row does not have a parent to hook to
            else if (entriesForExport[i].isSubType && !parentId) {
                response.error = true;
                response.message = "can't add a child type when there is no corresponding parent type";
                log.errorCount++;

            // send to Spira and update the response object
            } else {
                var sentToSpira = manageSendingToSpira ( entriesForExport[i], parentId, artifact, model.user, model.currentProject.id, fields, fieldType );

                parentId = sentToSpira.parentId;
                response.details = sentToSpira;

                // handle success and error cases
                if (sentToSpira.error) {
                    log.errorCount++;
                    response.error = true;
                    response.message = sentToSpira.message;

                    //Sets error HTML modal
                    htmlOutput = HtmlService.createHtmlOutput('<p ' + INLINE_STYLING + '>Error sending ' + (i + 1) + ' of ' + (entriesForExport.length) + '</p>').setWidth(250).setHeight(75);
                    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Progress');

                } else {
                    log.successCount++;
                    response.newId = sentToSpira.newId;

                    //modal that displays the status of each artifact sent
                    htmlOutputSuccess = HtmlService.createHtmlOutput('<p ' + INLINE_STYLING + '>' + (i + 1) + ' of ' + (entriesForExport.length) + ' sent!</p>').setWidth(250).setHeight(75);
                    SpreadsheetApp.getUi().showModalDialog(htmlOutputSuccess, 'Progress');
                }
            }

            log.entries.push(response);
        }

        // review all activity and set final status
        log.status = log.errorCount ? (log.errorCount == log.entriesLength ? STATUS_ENUM.allError : STATUS_ENUM.someError) : STATUS_ENUM.allSuccess;

        // 5. SET MESSAGES AND FORMATTING ON SHEET
        var bgColors = [],
            notes = [],
            values = [];
        // first handle cell formatting
        for (var row = 0; row < sheetData.length; row++) {
            var rowBgColors = [],
                rowNotes = [],
                rowValues = [];
            for (var col = 0; col < fields.length; col++) {
                var bgColor = setFeedbackBgColor(sheetData[row][col], log.entries[row].error, fields[col], fieldType, artifact, model.colors ),
                    note = setFeedbackNote(sheetData[row][col], log.entries[row].error, fields[col], fieldType, log.entries[row].message ),
                    value = setFeedbackValue(sheetData[row][col], log.entries[row].error, fields[col], fieldType, log.entries[row].newId || "", log.entries[row].details.entry.isSubType );

                rowBgColors.push(bgColor);
                rowNotes.push(note);
                rowValues.push(value);
            }
            bgColors.push(rowBgColors);
            notes.push(rowNotes);
            values.push(rowValues);
        }
        sheetRange.setBackgrounds(bgColors).setNotes(notes).setValues(values);

        //return {log: log, fields: fields, fieldType: fieldType, cellText: cellText, bgColors: bgColors, sheetData: sheetData, notes: notes};
        return log;

    }
}



// function that reviews a specific cell against it's field and errors for providing UI feedback on errors
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldType - enum information about field types
// @param: artifact - the currently selected artifact
// @param: colors - object of colors to use based on different conditions
function setFeedbackBgColor (cell, error, field, fieldType, artifact, colors) {
    if (error) {
        // if we have a validation error, we can highlight the relevant cells if the art has no sub type
        if (!artifact.hasSubType) {
            if (field.required && cell === "") {
                return colors.warning;
            } else {
                // keep original formatting
                if (field.type == fieldType.subId || field.type == fieldType.id || field.unsupported) {
                    return colors.bgReadOnly;
                } else {
                    return null;
                }
            }

        // otherwise highlight the whole row as we don't know the cause of the problem
        } else {
            return colors.warning;
        }

    // no errors
    } else {
        // keep original formatting
        if (field.type == fieldType.subId || field.type == fieldType.id || field.unsupported) {
            return colors.bgReadOnly;
        } else {
            return null;
        }
    }
}



// function that reviews a specific cell against it's field and sets any notes required
// currently only adds error message as note to ID field
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldType - enum information about field types
// @param: message - relevant error message from the entry for this row
function setFeedbackNote (cell, error, field, fieldType, message) {
    // handle entries with errors - add error notes into ID field
    if (error && field.type == fieldType.id) {
        return message;
    } else {
        return null;
    }
}



function setFeedbackValue (cell, error, field, fieldType, newId, isSubType) {
    // when there is an error we don't change any of the cell data
    if (error) {
        return cell;

        // handle successful entries - ie add ids into right place
    } else {
        var newIdToEnter =  newId || "";
        if (!isSubType && field.type == fieldType.id) {
            return newIdToEnter;
        } else if (isSubType && field.type == fieldType.subId) {
            return newIdToEnter;
        } else {
            return cell;
        }
    }
}



// on determining that an entry should be sent to Spira, this function handles calling the API function, and parses the data on both success and failure
// @param: entry - object of the specific entry in format ready to attach to body of API request
// @param: parentId - int of the parent id for this specific loop - used for attaching subtype children to the right parent artifact
// @param: artifact - object of the artifact being used here to help manage what specific API call to use
// @param: user - user object for API call authentication
// @param: projectId - int of project id for API call
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldType - object of all field types with enums
function manageSendingToSpira (entry, parentId, artifact, user, projectId, fields, fieldType) {
    var data,
        output = {},
        // make sure correct artifact ID is sent to handler (ie type vs subtype)
        artifactIdToSend = entry.isSubType ? artifact.subTypeId : artifact.id,
        // only send a parentId value when dealing with subtypes
        parentIdToSend = entry.isSubType ? parentId : null;

    // set output parent id here so we know this function will always return a value for this
    output.parentId = parentId; 
    
    // send object to relevant artifact post service
    data = postArtifactToSpira ( entry, user, projectId, artifactIdToSend, parentIdToSend );

    // save data for logging to client
    output.entry = entry;
    output.httpCode = data.getResponseCode();
    output.artifact = {
        artifactId: artifactIdToSend,
        artifactObject: artifact
    };

    // parse the data if we have a success
    if (output.httpCode == 200) {
        output.fromSpira = JSON.parse(data.getContentText());

        // get the id/subType id of the newly created artifact
        var artifactIdField = getIdFieldName(fields, fieldType, entry.isSubType);
        output.newId = output.fromSpira[artifactIdField];


        // update the output parent ID to the new id only if the artifact has a subtype and this entry is NOT a subtype
        if (artifact.hasSubType && !entry.isSubType) {
            output.parentId = output.newId;
        }

    } else {
        //we have an error - so set the flag and the message
        output.error = true;
        if (data && data.getContentText()) {
            output.errorMessage = data.getContentText();
        } else {
            output.errorMessage = "send attempt failed";
        }

        // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
        if (artifact.hasSubType && !entry.isSubType) {
            output.parentId = 0;
        }
    }
    return output;
}



// returns an int of the total number of required fields for the passed in artifact
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
// @param: forSubType - bool to determine whether to check for sub type required fields (true), or not - defaults to false
function countRequiredFieldsByType (fields, forSubType) {
    var count = 0;
    for (var i = 0; i < fields.length; i++) {
        if (forSubType != "undefined" && forSubType) {
            if (fields[i].requiredForSubType) {
                count++;
            }
        } else if (fields[i].required) {
            count++;
        }
    }
    return count;
}



// check to see if a row of data has entries for all required fields
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
function rowCountRequiredFieldsByType (row, fields, forSubType) {
    var count = 0;
    for (var i = 0; i < row.length; i++) {
        if (forSubType != "undefined" && forSubType) {
            if (fields[i].requiredForSubType && !row[i]) {
                count++;
            }
        } else if (fields[i].required && !row[i]) {
            count++;
        }

    }
    return count;
}



// check to see if a row for an artifact with a subtype has a field that can't be present if subtype fields are filled in
// this can be useful to make sure that one field - eg Test Case Name would make sure a test step is not created to avoid any confusion
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
function rowBlocksSubType (row, fields) {
    var result = false;
    for (var column = 0; column < row.length; column++) {
        if (fields[column].forbidOnSubType && row[column]) {
            result = true;
        }
    }
    return result;
}



// checks to see if the row is valid - ie required fields present and correct as expected
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: rowChecks - object with different properties for different checks required
function rowHasProblems (rowChecks) {
    var problems = null;
    if (!rowChecks.hasSubType && rowChecks.countRequiredFields < rowChecks.totalFieldsRequired) {
        problems = "Fill in all required fields";
    } else if (rowChecks.hasSubType) {
        if (rowChecks.countSubTypeRequiredFields < rowChecks.totalSubTypeFieldsRequired && !rowChecks.countRequiredFields) {
            problems = "Fill in all required fields";
        } else if (rowChecks.countRequiredFields < rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
            problems = "Fill in all required fields";
        } else if (rowChecks.countRequiredFields && (rowChecks.countRequiredFields == rowChecks.totalFieldsRequired || rowChecks.subTypeIsBlocked) ){
            problems = "It is unclear what artifact this is intended to be";
        }
    }
    return problems;
}



// based on field type and conditions, determines what fields are required for a given row
// e.g. all fields is default and standard, if a subtype is present (eg test step) - should it send only the main type or the sub type fields
// returns a int representing the relevant enum value
// @ param: rowChecks - object with different properties for different checks required
function relevantFields (rowChecks) {
    var fields = FIELD_MANAGEMENT_ENUMS.all;
    if (rowChecks.hasSubType) {
        if (rowChecks.countRequiredFieldsFilled == rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
            fields = FIELD_MANAGEMENT_ENUMS.standard;
        } else if (rowChecks.countSubTypeRequiredFields == rowChecks.totalSubTypeFieldsRequired && !(rowChecks.countRequiredFields == rowChecks.totalFieldsRequired || rowChecks.subTypeIsBlocked) ) {
            fields = FIELD_MANAGEMENT_ENUMS.subType;
        }
    }
    return fields;
}



// function creates a correctly formatted artifact object ready to send to Spira
// it works through each field type to validate and parse the values so object is in correct form
// any field that does not pass validation receives a null value
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: model - full model with info about fields, dropdowns, users, etc
// @param: fieldType - object of all field types with enums
// @param: artifactIsHierarchical - bool to tell function if this artifact has hierarchy (eg RQ and RL)
// @param: lastIndentPosition - int used for calculating relative indents for hierarchical artifacts
// @param: fieldsToFilter - enum used for selecting fields to not add to object - defaults to using all if omitted
function createEntryFromRow (row, model, fieldType, artifactIsHierarchical, lastIndentPosition, fieldsToFilter) {
    //create empty 'entry' object - include custom properties array here to avoid it being undefined later if needed
    var entry = {
            "CustomProperties": []
        },
        fields = model.fields;

    //we need to turn an array of values in the row into a validated object
    for (var index = 0; index < row.length; index++) {

        // first ignore entry that does not match the requirement specified in the fieldsToFilter
        if (fieldsToFilter == FIELD_MANAGEMENT_ENUMS.standard && fields[index].isSubTypeField ) {
            // skip the field
        } else if (fieldsToFilter == FIELD_MANAGEMENT_ENUMS.subType && !(fields[index].isSubTypeField || fields[index].isTypeAndSubTypeField) ) {
            // skip the field

        // in all other cases add the field
        } else {
            var value = null,
                customType = "",
                idFromName = 0;

            // double check data validation, convert dropdowns to required int values
            // sets both the value, and custom types - so that custom fields are handled correctly
            switch (fields[index].type) {

                // ID fields: restricted to numbers and blank on push, otherwise put
                case fieldType.id:
                case fieldType.subId:

                    customType = "IntegerValue";
                    break;

                // INT and NUM fields are both treated by Sheets as numbers
                case fieldType.int:
                case fieldType.num:
                    // only set the value if a number has been returned
                    if (!isNaN(row[index])) {
                        value = row[index];
                        customType = "IntegerValue";
                    }
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
                    }
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
                    idFromName = getIdFromName(row[index], fields[index].values);
                    if (idFromName) {
                        value = idFromName;
                        customType = "IntegerValue";
                    }
                    break;

                // MULTIDROPDOWNS - get id from relevant name, if one is present, set customtype to list value
                case fieldType.multi:
                    idFromName = getIdFromName(row[index], fields[index].values);
                    if (idFromName) {
                        value = idFromName;
                        customType = "IntegerListValue";
                    }
                    break;

                // USER fields - get id from relevant name, if one is present
                case fieldType.user:
                    idFromName = getIdFromName(row[index], model.projectUsers);
                    if (idFromName) {
                        value = idFromName;
                        customType = "IntegerValue";
                    }
                    break;

                // COMPONENT fields - get id from relevant name, if one is present
                case fieldType.component:
                    idFromName = getIdFromName(row[index], model.projectComponents);
                    if (idFromName) {
                        value = idFromName;
                        customType = "IntegerValue";
                    }
                    break;

                // RELEASE fields - get id from relevant name, if one is present
                case fieldType.release:
                    idFromName = getIdFromName(row[index], model.projectReleases);
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


            // HIERARCHICAL ARTIFACTS:
            // handle hierarchy fields - if required: checks artifact type is hierarchical and if this field sets hierarchy
            if (artifactIsHierarchical && fields[index].setsHierarchy) {
                // first get the number of indent characters
                var indentCount = countIndentCharacaters(value, model.indentCharacter);
                var indentPosition = setRelativePosition(indentCount, lastIndentPosition);

                // make sure to slice off the indent characters from the front
                // TODO should also trim white space at start
                value = value.slice(indentCount, value.length);

                // set the indent position for this row
                entry.indentPosition = indentPosition;
            }

            // CUSTOM FIELDS:
            // check whether field is marked as a custom field and as the required property number
            if (fields[index].isCustom && fields[index].propertyNumber) {

                // if field has data create the object
                if (value) {
                    var customObject = {};
                    customObject.PropertyNumber = fields[index].propertyNumber;
                    customObject[customType] = value;

                    entry.CustomProperties.push(customObject);
                }

            // STANDARD FIELDS:
            // add standard fields in standard way - only add if field contains data
            } else if (value) {
                entry[fields[index].field] = value;
            }
        }

    }

    return entry;
}



// find the corresponding ID for a string value - eg from a dropdown
// dropdowns can only contain one item per row so we have to now get the IDs for sending to Spira
// @param: string - the string of the name value specified
// @param: list - the array of items with keys for id and name values
function getIdFromName (string, list) {
    for (var i = 0; i < list.length; i++) {
        if (list[i].name == string) {
            return list[i].id;
        }
    }
    // return 0 if there's no match
    return 0;
}


// finds and returns the field name for the specific artifiact's ID field
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldType - object of all field types with enums
// @param: getSubType - optioanl bool to specify to return the subtype Id field, not the normal field (where two exist)
function getIdFieldName (fields, fieldType, getSubType) {
    for (var i = 0; i < fields.length; i++) {
        var fieldToLookup = getSubType ? "subId" : "id";
        if (fields[i].type == fieldType[fieldToLookup]) {
            return fields[i].field;
        }
    }
    return null;
}


// returns the count of the number of indent characters and returns the value
// @param: field - a single field string - one already designated as containing hierarchy info
// @param: indentCharacter - the character used to denote an indent - e.g. ">"
function countIndentCharacaters (field, indentCharacter) {
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



// returns the correct relative indent position - based on the previous relative indent and other logic (int neg, pos, or zero)
// the first time this is called, last position will be null
// setting indent to -10 is a hack to push the first item (hopefully) all the way to the root position - ie ignore any indents placed by user on first item
// Currently the API does not support a call to place an artifact at a certain location.
// @param: indentCount - int of the number of indent characters set by user
// @param: lastIndentPosition - int of the actual indent position used for the preceding entry/row
function setRelativePosition (indentCount, lastIndentPosition) {
    return (lastIndentPosition === null) ? -10 : indentCount - lastIndentPosition;
}