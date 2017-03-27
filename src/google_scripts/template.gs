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
  stdColorRange.setBackground('#073642');
  stdColorRange.setFontColor('#fff');
  var cusColorRange = sheet.getRange(data.requirements.customRange);
  cusColorRange.setBackground('#1398b9');
  cusColorRange.setFontColor('#fff');

  //get range for requirement numbers and set as greyed out
  var reqIdRange = sheet.getRange('A3:A200');
  reqIdRange.setBackground('#a6a6a6');

  //set customfield cells as grey if inactive
  var customCellRange = sheet.getRange('N3:AQ200');
  customCellRange.setBackground('#a6a6a6');

  //set column A to present a warning if the user trys to write in a value
  var protection = reqIdRange.protect().setDescription('Exported items must not have a requirement number');
  //set warning. Remove this to make the column un-writable
  protection.setWarningOnly(true);

  sheet.getRange('A1:M1').merge().setValue("Requirements Standard Fields").setHorizontalAlignment("center");
  sheet.getRange('N1:AQ1').merge().setValue("Custom Fields").setHorizontalAlignment("center");

  //append headings to sheet
  sheet.appendRow(data.requirements.headings)

  //set custom headings if they exist
  //pass in custom field range, data model, and custom column to be used for background coloring
  customHeadSetter(sheet.getRange('N2:AQ2'), data, sheet.getRange('N3:N200'));

  //loop through model sizes data and set columns to correct width
  for(var i = 0; i < data.requirements.sizes.length; i++){
    sheet.setColumnWidth(data.requirements.sizes[i][0],data.requirements.sizes[i][1]);
  }

  //custom field validation and dropdowns
  customContentSetter(sheet.getRange(data.requirements.customCellRange), data)

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

function customHeadSetter(range, data, col){

  //shorten variable
  var fields = data.requirements.customFields
  //loop through model custom fields data
  //take passed in range and only overwrite the fields if a value is present in the model
  for(var i = 0; i < fields.length; i++){
    //get cell and offset by one column very iteration
    var cell = range.getCell(1, i + 1)
    //set heading
    cell.setValue('Custom Field ' + (i + 1) + '\n' + fields[i].Name).setWrap(true);
    //get column and offset every iteration and set background
    var column = col.offset(0, i)
    column.setBackground('#fff');
  }
}

function customContentSetter(range, data){
  //shorten variable
  var customs = data.requirements.customFields;
  //grab users list from owners dropdown
  var users = data.requirements.dropdowns['Owner'];
  //loop through custom property fields
  for(var i = 0; i < customs.length; i++){
    //check if field matches {2: integer} or {3: float}
    if(customs[i].CustomPropertyTypeId == 2 || customs[i].CustomPropertyTypeId == 3){
      //get first cell in range
      var cell = range.getCell(1, i + 1);
      //get column range (x : x)
      var column = columnRanger(cell);
      //set number only data validation
      var rule = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).setAllowInvalid(false).setHelpText('Must be a positive integer').build();
      column.setDataValidation(rule);
    }
    //boolean
    if(customs[i].CustomPropertyTypeId == 4){
      var list = ["Yes", "No"];
      var cell = range.getCell(1, i + 1);
      var cellsTop = cell.getA1Notation();
      var cellsEnd = cell.offset(200, 0).getA1Notation();
      var column = SpreadsheetApp.getActive().getRange(cellsTop + ':' + cellsEnd);

     var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
     column.setDataValidation(rule);
    }
    //date
    if(customs[i].CustomPropertyTypeId == 5){
      var cell = range.getCell(1, i + 1);
      var column = columnRanger(cell);
      //set number only data validation
      var rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Must be a valid date').build();
      column.setDataValidation(rule);
    }
    //List and MultiList
    if(customs[i].CustomPropertyTypeId == 6 || customs[i].CustomPropertyTypeId == 7){
      var list = [];
      for(var j = 0; j < customs[i].CustomList.Values.length; j++){
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
    if(customs[i].CustomPropertyTypeId == 8){
      var list = [];
      var len = users.length;
      for(var j = 0; j < len; j++){
        list.push(users[j]);
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

function columnRanger(cell){
  //get the column
  var col = cell.getColumn();
  //get column letter
  col = columnToLetter(col);
  //get column range for data validation
  var column = SpreadsheetApp.getActive().getRange(col + ':' + col);

  return column;
}

//open source column to letter function **Adam L from Stack OverFlow
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}



