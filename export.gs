/*

Export functions

The main function takes the entire data model and the artifact type and calls the child function to set various object values before sending the finished objects to SpiraTeam

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
