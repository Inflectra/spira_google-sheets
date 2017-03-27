//import function, basic for now
//stretch goal is to have this as a useful import
function importer(currentUser){
  //get spreadsheet and active first sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheets()[0];

  // needed for eventual actual importer
  // var paramsCount = '/services/v5_0/RestService.svc/projects/1/requirements/count?username=';
  // var count = getFetch(currentUser, paramsCount )

  //call defined fetch function
  //current params has count set to 35, this can be set/changed programmatically with the count call listed above (stretch goal)
  var params = '/services/v5_0/RestService.svc/projects/1/requirements?starting_row=1&number_of_rows=35&username=';
  var data = fetcher(currentUser, params)

  //get first row range
  var range = sheet.getRange("A3:AQ3");

  //loop through cells in range
  for(var i = 0; i < data.length; i++){
    var ss_i = i + 1
    range.getCell(ss_i, 1).setValue(data[i].RequirementId);
    range.getCell(ss_i, 2).setValue(data[i].Name);
    range.getCell(ss_i, 3).setValue(data[i].Description);
    range.getCell(ss_i, 4).setValue(data[i].ReleaseVersionNumber);
    range.getCell(ss_i, 5).setValue(data[i].RequirementTypeName);
    range.getCell(ss_i, 6).setValue(data[i].ImportanceName);
    range.getCell(ss_i, 7).setValue(data[i].StatusName);
    range.getCell(ss_i, 8).setValue(data[i].EstimatePoints);
    range.getCell(ss_i, 9).setValue(data[i].AuthorName);
    range.getCell(ss_i, 10).setValue(data[i].OwnerName);
    range.getCell(ss_i, 11).setValue(data[i].ComponentId);

    //moves the range down one row
    range = range.offset(1, 0, 43);
 }
}