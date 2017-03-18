//export function pulled from Code.gs
//takes item {cell}, list {array}, and isObj {bool}
//isObj is true if list is an object, i.e in the case of the users array
function mapper(item, list){
  var val = 1;
  for (var i = 0; i < list.length; i++){
    if (item == list[i][1]) {val = list[i][0]}
  }
  return val;
}

//function richData(data){
//  var textArr = data.split(' ');
//
//  for (var i = 0; i < textArr.length; i++){
//
//    var word = textArr[i]
//    //.isBold()
//
////    var italic = textArr[i].getFontStyle()
////    var underline = textArr[i].getFOntLines()
////
//    Logger.log(word)// + italic + underline
//
//  }
//  return data;
//}

function indenter(cell){
  // var indentCount = 0;
  // //check for indent character '>'
  // if(cell && cell[0] === '>'){
  // //increment indent counter while there are '>'s present
  //   while (cell[0] === '>'){
  //     //get entry length for slice
  //     var len = cell.length;
  //     //slice the first character off of the entry
  //     cell = cell.slice(1, len);
  //     indentCount++;
  //   }
  //   xObj['IndentLevel'] = 'AAB';
  // }
  return 'AAA'
}

function exporter(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];
  //number of cells in a row
  var range = sheet.getRange(data.templateData.requirements.cellRange)
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

    //loop through cells in row
    for (var i = 0; i < reqs.JSON_headings.length; i++){

      //get cell value
      var cell = range.offset(j, i).getValue();

      //passes description data to richData function to attach HTML tags for spirateam
      //if(i === 2.0){ cell = richData(cell) }

      //shorten variables
      var users = data.userData.projUserWNum;

      //pass values to mapper function
      //mapper iterates and assigns the values number based on the list order
      if(i === 4.0){ cell = mapper(cell, reqs.dropdowns['Type']) }

      if(i === 5.0){ xObj['ImportanceId'] = mapper(cell, reqs.dropdowns['Importance']) }

      if(i === 6.0){ xObj['StatusId'] = mapper(cell, reqs.dropdowns['Status']) }

      if (i === 8.0){ xObj['AuthorId'] = mapper(cell, users) }

      if (i === 9.0){ xObj['OwnerId'] = mapper(cell, users) }



      //call indent checker and set indent amount
      xObj['IndentLevel'] = indenter();

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
// var responses = []
// for(var i = 0; i < bodyArr.length; i++){
//  //stringify
//  var JSON_body = JSON.stringify( bodyArr[i] );
//  //send JSON to export function
//  var response = requirementExportCall( JSON_body, data.templateData.currentProjectNumber, data.userData.currentUser )
//  //push API approval into array
//  responses.push(response.RequirementId)
// }




  //return responses
  return bodyArr
  //return JSON.stringify( bodyArr )
  //return JSON_body;
}

function requirementExportCall(body, projNum, currentUser){
  //unique url for requirement POST
  var params = '/services/v5_0/RestService.svc/projects/' + projNum + '/requirements?username=';
  //POST headers
  var init = {
   'method' : 'post',
   'contentType': 'application/json',
   'payload' : body
  };
  //call fetch with POST request
  var res = fetcher(currentUser, params, init);

  return res;
}