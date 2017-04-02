//import function, basic for now
//stretch goal is to have this as a useful import
function importer(currentUser, data) {
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
    for (var i = 0; i < data.length; i++) {
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
