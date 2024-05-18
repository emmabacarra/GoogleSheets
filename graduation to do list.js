// Adding a Timestamp to New Tasks

function addTimestamp(event){ 
  var sheet = event.source.getActiveSheet();
  var range = event.source.getActiveRange();

  if (event.oldValue == null && sheet.getSheetValues(1, range.getColumn(), 1, 1)[0] == "Remaining Tasks") {
    
    // Adding static date to date assigned
    range.offset(0,-2).setValue(new Date());

    // Adding static date to deadline (as default, and to force user to insert Date Assigned if necessary)
    range.offset(0,1).setValue(new Date());

    // Adding default progress to "New"
    range.offset(0,2).setValue("New");
  }
}


// Completed Tasks Actions

function migratingTasks(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;


  // Migrating to Completed Tasks

  // Check if the edit is in the 'Remaining Tasks' tab (Status)
  if (sheet.getName() == "Remaining Tasks" && e.value == "Completed") {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "Completed Tasks" tab
    var archiveSheet = e.source.getSheetByName("Completed Tasks");

    // Find the row index where "Date Assigned" is in column A, row 1 on "Completed Tasks"
    var sigmaRowIndex = archiveSheet.getRange("A1:A").createTextFinder("Date Assigned").findNext().getRow();

    // Insert a new row below "sigma"
    archiveSheet.insertRowAfter(sigmaRowIndex);

    // Copy the formatting and values into the new row below "Date Assigned"
    sourceRow.copyTo(archiveSheet.getRange(sigmaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(archiveSheet.getRange(sigmaRowIndex + 1, 1));

    // Set the fill color to #ffffff (white)
    //archiveSheet.getRange(sigmaRowIndex + 1, 1, 1, sourceRow.getNumColumns()).setBackground("#ffffff");

    // Set row height to 21
    archiveSheet.setRowHeight(sigmaRowIndex + 1, 21);

    // Delete the original row from the "Remaining Tasks" tab
    sheet.deleteRow(range.getRow());
  } 

  // Moving back to 'Remaining Tasks' if marked as 'On Going' or 'In Progress' in Completed Tasks
  if (sheet.getName() == "Completed Tasks" && (e.value == "On Going" || e.value == "In Progress" || e.value == "Pending Review")) {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "To Dos - Completed Tasks" tab
    var doingTasksSheet = e.source.getSheetByName("Remaining Tasks");

    // Find the row index where "Example Task" is in column C, row 2 on "Remaining Tasks" tab
    var alphaRowIndex = doingTasksSheet.getRange("C2:C").createTextFinder("Example Task").findNext().getRow();

    // Insert a new row below "beta"
    doingTasksSheet.insertRowAfter(alphaRowIndex);

    // Copy the formatting and values into the new row below "Example Task"
    sourceRow.copyTo(doingTasksSheet.getRange(alphaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(doingTasksSheet.getRange(alphaRowIndex + 1, 1));

    // Set the fill color to #ffffff for columns C to J
    // doingTasksSheet.getRange(alphaRowIndex + 1, 3, 1, sourceRow.getNumColumns() - 2).setBackground("#ffffff");

    // Set the row height to 21
    doingTasksSheet.setRowHeight(alphaRowIndex + 1, 21);

    // Delete the original row from the "Archived Tasks" tab
    sheet.deleteRow(range.getRow());
  }
}













