//export function pulled from Code.gs
//takes item {cell}, list {array}, and isObj {bool}
//isObj is true if list is an object, i.e in the case of the users array

function exporter(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];
  //number of cells in a row
  var range = sheet.getRange(data.templateData.requirements.cellRange)
  var customRange = sheet.getRange(data.templateData.requirements.customCellRange)
  //var i
  var isRowEmpty = false;
  var numberOfRows = 0;
  var row = 0;
  var col = 0;
  var bodyArr = [];

  //shorten variable
  var reqs = data.templateData.requirements;

  //loop through and collect number of rows that contain data
  while (isRowEmpty === false){
    //select row
    var newRange = range.offset(row, col, reqs.cellRangeLength);
    //check if the row is empty
    if ( newRange.isBlank() ){
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
  for (var j = 0; j < numberOfRows + 1; j++){

    //initialize/clear new object for row values
    var xObj = {}

    //send current row
    var row = customRange.offset(j, 0)
    xObj['CustomProperties'] = customBuilder(data, row)


    //loop through cells in row
    for (var i = 0; i < reqs.JSON_headings.length; i++){


      //get cell value
      var cell = range.offset(j, i).getValue();

      //get cell Range for req# insertion after export
      if(i=== 0.0) { xObj['idField'] = range.offset(j, i).getCell(1, 1) }

      //call indent checker and set indent amount
      if(i === 1.0) { xObj['indentCount'] = indenter(cell) }

      //shorten variables
      var users = data.userData.projUserWNum;

      //pass values to mapper function
      //mapper iterates and assigns the values number based on the list order
      if(i === 4.0){ cell = mapper(cell, reqs.dropdowns['Type']) }

      if(i === 5.0){ xObj['ImportanceId'] = mapper(cell, reqs.dropdowns['Importance']) }

      if(i === 6.0){ xObj['StatusId'] = mapper(cell, reqs.dropdowns['Status']) }

      if (i === 8.0){ xObj['AuthorId'] = mapper(cell, users) }

      if (i === 9.0){ xObj['OwnerId'] = mapper(cell, users) }


      //if empty add null otherwise add the cell
      // ...to the object under the proper key relative to its location on the template
      //Offset by 2 for proj name and indent level
      if (cell === ""){
        xObj[reqs.JSON_headings[i]] = null;
      } else {
        xObj[reqs.JSON_headings[i]] = cell;
      }

    }

    //if not empty add object
    //entry MUST have a name
    if ( xObj.Name ) {
      xObj['ProjectName'] = data.templateData.currentProjectName;
      bodyArr.push(xObj)
    }

  }

  // set up to individually add each requirement to spirateam
  // maybe there's a way to bulk add them instead of individual calls?
 var responses = []
 var test = '';
 for(var i = 0; i < bodyArr.length; i++){
  //stringify
  var JSON_body = JSON.stringify( bodyArr[i] );
  //send JSON to export function
   var response = new Promise(function(res, rej){
     res( requirementExportCall( JSON_body, data.templateData.currentProjectNumber, data.userData.currentUser ) );
   });

   response.then(function(resp){
     responses.push(resp.RequirementId)
     //set returned ID
     bodyArr[i].idField.setValue('RQ:' + resp.RequirementId)
     if (bodyArr[i].indentCount > 0){
      var indentNum = bodyArr[i].indentCount;
      requirementIndentCall(data.templateData.currentProjectNumber, data.userData.currentUser, resp.RequirementId, indentNum)
    }
   });

  //push API approval into array
  //responses.push(response.RequirementId)

 }




  return responses
  //return test
}

function requirementExportCall(body, projNum, currentUser, indent){
  //unique url for requirement POST
  var params = '/services/v5_0/RestService.svc/projects/' + projNum + '/requirements?username=';
  //POST headers
  var init = {
   'method' : 'post',
   'contentType': 'application/json',
   'payload' : body
  };
  //call fetch with POST request
  return fetcher(currentUser, params, init);
}

function requirementIndentCall(projNum, currentUser, reqId, numOfIndents){
  //unique url for indent POST
  var params = '/services/v5_0/RestService.svc/projects/' + projNum + '/requirements/' + reqId + '/indent?username=';
  //POST headers
  var init = {
   'method' : 'post',
   'contentType': 'application/json',
  };
  for(var i = 1; i <= numOfIndents; i++){
    fetcher(currentUser, params, init);
  }
}

//map cell data to their corresponding IDs for export to spirateam
function mapper(item, list){
  //set return value to 1 on err
  var val = 1;
  //loop through model for variable being mapped
  for (var i = 0; i < list.length; i++){
    //cell value matches model value assign id number
    if (item == list[i][1]) {val = list[i][0]}
  }
  return val;
}

//gets full model data and custom properites cell range
function customBuilder(data, rowRange){
  //shorten variables
  var customs = data.templateData.requirements.customFields;
  var users = data.userData.projUserWNum;
  //length of custom data to optimise perf
  var len = customs.length;
  //custom props array of objects to be returned
  var customProps = [];
  //loop through cells based on custom data fields
  for(var i = 0; i < len; i++){
    //assign custom property to variable
    var customData = customs[i];
    //get cell data
    var cell = rowRange.offset(0, i).getValue()
    //check if the cell is empty
    if (cell !== ""){
      //call custom content function and push data into array from export
      customProps.push( customFiller(cell, customData, users) )
    }
  }
  //custom properties array ready for API export
  return customProps
}

//gets specific cell and custom property data for that column
function customFiller(cell, data, users){
  //all custom values need a property number
  //set it and add to object for return
  var propNum = data.PropertyNumber;
  var prop = {PropertyNumber: propNum}

 //check data type of custom fields and assign values if condition is met
 if(data.CustomPropertyTypeName == 'Text'){
   prop['StringValue'] = cell;
 }

 if(data.CustomPropertyTypeName == 'Integer'){
   prop['IntegerValue'] = cell;
 }

 if(data.CustomPropertyTypeName == 'Decimal'){
   prop['DecimalValue'] = cell;
 }

 if(data.CustomPropertyTypeName == 'Boolean'){
   cell == "Yes" ? prop['BooleanValue'] = true : prop['BooleanValue'] = false;
 }

 if(data.CustomPropertyTypeName == 'List'){
   var len = data.CustomList.Values.length;
   //loop through custom list and match name to cell value
   for (var i = 0; i < len; i++){
     if (cell == data.CustomList.Values[i].Name){
       //assign list value number to integer
       prop['IntegerValue'] = data.CustomList.Values[i].CustomPropertyValueId
     }
   }
 }

  if(data.CustomPropertyTypeName == 'Date'){
    //parse date into milliseconds
    cell = Date.parse(cell);
    //concat values accepted by spira and assign to correct prop
    prop['DateTimeValue'] = "\/Date(" + cell + ")\/";
  }


 if(data.CustomPropertyTypeName == 'MultiList'){}

 if(data.CustomPropertyTypeName == 'User'){
   var len = users.length
   for (var i = 0; i < len; i++){
     if (cell == users[i][1]){
       prop['IntegerValue'] = users[i][0];
     }
   }
 }


  return prop;
}

function indenter(cell){
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
  return indentCount
}