/*

Template creation function (.gs)

This function creates a template based on the model template data (TODO: currently only creates requirements template)

@param {object}: data model
*/


//function for template creation
function templateLoader(data) {
    //call clear function and clear spreadsheet depending on user input
    clearAll();

    //select open file and select first tab
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getSheets()[0];

    //shorten variable
    var dropdownColumnAssignments = data.requirements.dropdownColumnAssignments;

    //set sheet name to model name
    sheet.setName(data.currentProjectName + ' - ' + data.currentArtifactName);

    //set heading colors and font colors for standard and custom ranges
    var stdColorRange = sheet.getRange(data.requirements.standardRange);
    stdColorRange.setBackground('#073642');
    stdColorRange.setFontColor('#fff');
    var cusColorRange = sheet.getRange(data.requirements.customRange);
    cusColorRange.setBackground('#1398b9');
    cusColorRange.setFontColor('#fff');

    //get range for requirement numbers and set color
    //color set to grey to denote unwritable field
    var reqIdRange = sheet.getRange('A3:A200');
    reqIdRange.setBackground('#a6a6a6');

    //set customfield cells as grey if inactive
    var customCellRange = sheet.getRange('N3:AQ200');
    customCellRange.setBackground('#a6a6a6');

    //set column A to present a warning if the user tries to write in a value
    var protection = reqIdRange.protect().setDescription('Exported items must not have a requirement number');
    //set warning. Remove this to make the column un-writable
    protection.setWarningOnly(true);

    //set title range and center
    sheet.getRange(data.requirements.standardTitleRange).merge().setValue("Requirements Standard Fields").setHorizontalAlignment("center");
    sheet.getRange(data.requirements.customTitleRange).merge().setValue("Custom Fields").setHorizontalAlignment("center");

    //append headings to sheet
    sheet.appendRow(data.requirements.headings)

    //set custom headings if they exist
    //pass in custom field range, data model, and custom column to be used for background coloring
    customHeadSetter(sheet.getRange(data.requirements.customHeaders), data, sheet.getRange(data.requirements.customColumnLength));

    //loop through model sizes data and set columns to correct width
    for (var i = 0; i < data.requirements.sizes.length; i++) {
        sheet.setColumnWidth(data.requirements.sizes[i][0], data.requirements.sizes[i][1]);
    }

    //main custom field function assigns type, dropdowns, datavalidation etc. See function for details.
    customContentSetter(sheet.getRange(data.requirements.customCellRange), data)

    //loop through dropdowns model data
    for (var i = 0; i < dropdownColumnAssignments.length; i++) {
        //variable assignment from dropdown object
        var letter = dropdownColumnAssignments[i][1];
        var name = dropdownColumnAssignments[i][0];
        //array that will hold dropdown values
        var list = [];
        //loop through 2D arrays and form standard array
        for (var j = 0; j < data.requirements.dropdowns[name].length; j++) {
            list.push(data.requirements.dropdowns[name][j][1])
        }

        //set range to entire column excluding top two rows (offset)
        var cell = SpreadsheetApp.getActive().getRange(letter + ':' + letter).offset(2, 0);
        //require list of values as dropdown and entered values
        //require value in list: list variable is from the model, true shows dropdown arrow
        //allow invalid set to false does not allow invalid entries
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
        cell.setDataValidation(rule);
    }
    //loop through data model
    //set 'number only' columns to only accept numbers
    for (var i = 0; i < data.requirements.requireNumberFields.length; i++) {
        var colLetter = data.requirements.requireNumberFields[i];
        var column = SpreadsheetApp.getActive().getRange(colLetter + ':' + colLetter);
        //does not allow negative numbers or non-integers
        //sets a tooltip explaining cell rules
        var rule = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).setAllowInvalid(false).setHelpText('Must be a positive integer').build();
        column.setDataValidation(rule);
    }
}

/*
Custom header setter function

@params {range}: google range value i.e (A1:B2)
@params {object}: template data model
@params {range}: google range value
*/

//Sets headings for custom fields
function customHeadSetter(range, data, col) {

    //shorten variable
    var fields = data.requirements.customFields

    //loop through model custom fields data
    //take passed in range and only overwrite the fields if a value is present in the model
    for (var i = 0; i < fields.length; i++) {
        //get cell and offset by one column very iteration
        var cell = range.getCell(1, i + 1)
            //set heading
        cell.setValue('Custom Field ' + (i + 1) + '\n' + fields[i].Name).setWrap(true);
        //get column and offset every iteration and set background
        var column = col.offset(0, i)
        column.setBackground('#fff');
    }
}

/*
Custom content setter function

@params {range}: google range value i.e (A1:B2)
@params {object}: template data model
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
            var column = columnRanger(cell);

            //set number only data validation
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
            var column = columnRanger(cell);
            //set number only data validation
            var rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Must be a valid date').build();
            column.setDataValidation(rule);
        }
        //List and MultiList
        if (customs[i].CustomPropertyTypeId == 6 || customs[i].CustomPropertyTypeId == 7) {
            var list = [];
            for (var j = 0; j < customs[i].CustomList.Values.length; j++) {
                list.push(customs[i].CustomList.Values[j].Name);
            }
            var cell = range.getCell(1, i + 1);
            var cellsTop = cell.getA1Notation();
            var cellsEnd = cell.offset(200, 0).getA1Notation();
            var column = SpreadsheetApp.getActive().getRange(cellsTop + ':' + cellsEnd);

            var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
            column.setDataValidation(rule);
        }
        //users
        if (customs[i].CustomPropertyTypeId == 8) {
            var list = [];
            var len = users.length;
            for (var j = 0; j < len; j++) {
                list.push(users[j][1]);
            }
            var cell = range.getCell(1, i + 1);
            var cellsTop = cell.getA1Notation();
            var cellsEnd = cell.offset(200, 0).getA1Notation();
            var column = SpreadsheetApp.getActive().getRange(cellsTop + ':' + cellsEnd);

            var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
            column.setDataValidation(rule);
        }
    }

}

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
