//export function pulled from Code.gs
//takes item {cell}, list {array}, and isObj {bool}
//isObj is true if list is an object, i.e in the case of the users array

function exporter(data) {

    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = spreadSheet.getSheets()[0];
    //number of cells in a row
    var range = sheet.getRange(data.templateData.requirements.cellRange)
    var customRange = sheet.getRange(data.templateData.requirements.customCellRange)
        //var i
    var isRowEmpty = false;
    var numberOfRows = 0;
    var row = 0;
    var col = 0;
    var bodyArr = [];
    var responses = [];
    var xObjArr = [];

    //shorten variable
    var reqs = data.templateData.requirements;

    var htmlOutput = HtmlService.createHtmlOutput('<p>Preparing your data for export!</p>').setWidth(250).setHeight(75);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Progress');

    //loop through and collect number of rows that contain data
    while (isRowEmpty === false) {
        //select row
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

    //loop through standard rows
    for (var j = 0; j < numberOfRows + 1; j++) {

        //initialize/clear new object for row values
        var xObj = {}

        //send current row
        var row = customRange.offset(j, 0)
        xObj['CustomProperties'] = customHeaderRowBuilder(data, row)
        xObj['positionNumber'] = 0;


        //loop through cells in row
        for (var i = 0; i < reqs.JSON_headings.length; i++) {


            //get cell value
            var cell = range.offset(j, i).getValue();

            //get cell Range for req# insertion after export
            if (i === 0.0) { xObj['idField'] = range.offset(j, i).getCell(1, 1) }

            //call indent checker and set indent amount
            if (i === 1.0) {
                //call indent function
                //counts the number of ">"s to assign an indent value
                xObj['indentCount'] = indenter(cell)

                //remove '>' symbols from requirement name string
                while (cell[0] == '>' || cell[0] == ' ') {
                    //removes first charactor if it's a space or ">"
                    cell = cell.slice(1)
                }
            }

            //shorten variables
            var users = data.userData.projUserWNum;

            //pass values to mapper function
            //mapper iterates and assigns the values number based on the list order

            //TODO add requirements and components
            if (i === 3.0) { xObj['ReleaseId'] = mapper(cell, reqs.dropdowns['Version Number']) }

            //The type property is currently bugged in the API
            //All are hard coded to id = 1 for 'feature'
            if (i === 4.0) { cell = mapper(cell, reqs.dropdowns['Type']) }

            if (i === 5.0) { xObj['ImportanceId'] = mapper(cell, reqs.dropdowns['Importance']) }

            if (i === 6.0) { xObj['StatusId'] = mapper(cell, reqs.dropdowns['Status']) }

            if (i === 8.0) { xObj['AuthorId'] = mapper(cell, users) }

            if (i === 9.0) { xObj['OwnerId'] = mapper(cell, users) }

            if (i === 10.0) { xObj['ComponentId'] = mapper(cell, reqs.dropdowns['Components']) }


            //if empty add null otherwise add the cell
            // ...to the object under the proper key relative to its location on the template
            //Offset by 2 for proj name and indent level
            if (cell === "") {
                xObj[reqs.JSON_headings[i]] = null;
            } else {
                xObj[reqs.JSON_headings[i]] = cell;
            }

        }

        //if not empty add object
        //entry MUST have a name
        if (xObj.Name) {
            xObj['ProjectName'] = data.templateData.currentProjectName;

            xObjArr.push(xObj);
        }

        xObjArr = parentChildSetter(xObjArr);
    }

    // set up to individually add each requirement to spirateam
    // maybe there's a way to bulk add them instead of individual calls?
    var testArr = []
    var isError = null;
    var errorLog = [];
    var len = xObjArr.length
    for (var i = 0; i < len; i++) {
        //stringify
        var JSON_body = JSON.stringify(xObjArr[i]);

        //send JSON to export function
        var response = requirementExportCall(JSON_body, data.templateData.currentProjectNumber, data.userData.currentUser, xObjArr[i].positionNumber);


        if (response.getResponseCode() === 200) {
            response = JSON.parse(response.getContentText())

            responses.push(response.RequirementId)
            //set returned ID
            xObjArr[i].idField.setValue(response.RequirementId)

            htmlOutputSuccess = HtmlService.createHtmlOutput('<p>' + (i + 1) + ' of ' + (len) + ' sent!</p>').setWidth(250).setHeight(75);
            SpreadsheetApp.getUi().showModalDialog(htmlOutputSuccess, 'Progress');
        } else {
            errorLog.push( response.getContentText());
            isError = true;
            //set returned ID
            //removed by request can be added back if wanted in future versions
            //xObjArr[i].idField.setValue('Error')

            //Sets error HTML
            htmlOutput = HtmlService.createHtmlOutput('<p>Error for ' + (i + 1)  + ' of ' + (len) + '</p>').setWidth(250).setHeight(75);
            SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Progress');
        }
    }
    //return the error flag and array with error text reponses
    return [isError, errorLog];
}

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

//gets full model data and custom properites cell range
function customHeaderRowBuilder(data, rowRange) {
    //shorten variables
    var customs = data.templateData.requirements.customFields;
    var users = data.userData.projUserWNum;
    //length of custom data to optimise perf
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
        cell = parseInt(cell);
        prop['IntegerValue'] = cell;
    }

    if (data.CustomPropertyTypeName == 'Decimal') {
        prop['DecimalValue'] = cell;
    }

    if (data.CustomPropertyTypeName == 'Boolean') {
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

        //single item exported
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
        var len = users.length
        for (var i = 0; i < len; i++) {
            if (cell == users[i][1]) {
                prop['IntegerValue'] = users[i][0];
            }
        }
    }


    return prop;
}

function indenter(cell) {
    var indentCount = 0;
    //check for indent character '>'
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
    var len = arr.length;
    //set to 2 to make the math work on the first indent level 2 - 1 = indent level 1
    var count = 0;

    for (var i = 0; i < len; i++) {
        if (arr[i].indentCount > 0 && arr[i].indentCount !== count) {
            arr[i].positionNumber = arr[i].indentCount - count;
            count = arr[i].indentCount;
        }

        if (arr[i].indentCount == 0) {
            arr[i].positionNumber = -10;
            //reset count variable
            count = 0;
        }
    }
    return arr;
}
