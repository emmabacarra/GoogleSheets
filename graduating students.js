// ------- This is the new version for activating with clickable button (using macros) -------
function updateSheets() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var graduatingStudentsSheet = ss.getSheetByName('Graduating Students');
  
  // Get the 'Form Responses' and 'Slide Submissions' sheets by their IDs
  var formResponsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RSVP Responses');
  var slideSubmissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Slide Submissions');

  // Add activation notification
  graduatingStudentsSheet.getRange('C1:F1').setValue('Please wait, update in progress');
  SpreadsheetApp.getActive().toast("Please wait, updating list with new responses");

  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var slideSubmissionsData = slideSubmissionsSheet.getDataRange().getValues();

  for (var i = 2; i < formResponsesData.length; i++) {
    var id = formResponsesData[i][4]; // Column E
    var tickets = formResponsesData[i][6]; // Column G

    // Find the rows in 'Graduating Students' that match the ID
    var graduatingStudentsData = graduatingStudentsSheet.getDataRange().getValues();
    for (var j = 2; j < graduatingStudentsData.length; j++) {
      if (graduatingStudentsData[j][2] === id) { // Column C
        // Copy the tickets value to column L
        graduatingStudentsSheet.getRange(j + 1, 12).setValue(tickets); // Column L

        // If tickets > 0, set column K to true
        if (tickets > 0) {
          graduatingStudentsSheet.getRange(j + 1, 11).setValue(true); // Column K
        }

        // Set column J to true
        graduatingStudentsSheet.getRange(j + 1, 10).setValue(true); // Column J
      }
    }
  }

  for (var i = 2; i < slideSubmissionsData.length; i++) {
    if (slideSubmissionsData[i][14] !== '') { // Column O
      var id = slideSubmissionsData[i][7]; // Column H

      // Find the rows in 'Graduating Students' that match the ID
      var graduatingStudentsData = graduatingStudentsSheet.getDataRange().getValues();
      for (var j = 2; j < graduatingStudentsData.length; j++) {
        if (graduatingStudentsData[j][2] === id) { // Column C
          // Set column M to true
          graduatingStudentsSheet.getRange(j + 1, 13).setValue(true); // Column M
        }
      }
    }
  }

  // Add completion notification
  SpreadsheetApp.getActive().toast("Update Complete");

  // adding timestamp for last update
  graduatingStudentsSheet.getRange('C1:F1').setValue('Last updated: ' + Utilities.formatDate(new Date(), "GMT-7", "MM/dd HH:mm"));
}



// ------- This is the old version for activating with checkbox -------
// function updateSheets() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var graduatingStudentsSheet = ss.getSheetByName('Graduating Students');
  
//   // Get the 'Form Responses' and 'Slide Submissions' sheets by their IDs
//   var formResponsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RSVP Responses');
//   var slideSubmissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Slide Submissions');

//   // Check if the value in A1 of the 'Graduating Students' sheet is true
//   if (graduatingStudentsSheet.getRange('A1').getValue() === true) {
//     // Set the value of the merged cell to "Please wait, updating list with new responses"
//     graduatingStudentsSheet.getRange('C1:F1').setValue('Please wait, updating list with new responses');

//     var formResponsesData = formResponsesSheet.getDataRange().getValues();
//     var slideSubmissionsData = slideSubmissionsSheet.getDataRange().getValues();

//     for (var i = 2; i < formResponsesData.length; i++) {
//       var email = formResponsesData[i][1]; // Column B
//       var tickets = formResponsesData[i][6]; // Column G

//       // Find the rows in 'Graduating Students' that match the email
//       var graduatingStudentsData = graduatingStudentsSheet.getDataRange().getValues();
//       for (var j = 2; j < graduatingStudentsData.length; j++) {
//         if (graduatingStudentsData[j][2] === email) { // Column C
//           // Copy the tickets value to column K
//           graduatingStudentsSheet.getRange(j + 1, 11).setValue(tickets); // Column K

//           // If tickets > 0, set column J to true
//           if (tickets > 0) {
//             graduatingStudentsSheet.getRange(j + 1, 10).setValue(true); // Column J
//           }

//           // Set column I to true
//           graduatingStudentsSheet.getRange(j + 1, 9).setValue(true); // Column I
//         }
//       }
//     }

//     for (var i = 2; i < slideSubmissionsData.length; i++) {
//       if (slideSubmissionsData[i][13] !== '') { // Column N
//         var email = slideSubmissionsData[i][1]; // Column B

//         // Find the rows in 'Graduating Students' that match the email
//         var graduatingStudentsData = graduatingStudentsSheet.getDataRange().getValues();
//         for (var j = 2; j < graduatingStudentsData.length; j++) {
//           if (graduatingStudentsData[j][2] === email) { // Column C
//             // Set column L to true
//             graduatingStudentsSheet.getRange(j + 1, 12).setValue(true); // Column L
//           }
//         }
//       }
//     }

//     // Clear the checked box
//     graduatingStudentsSheet.getRange('A1').setValue(false);

//     // adding timestamp for last update
//     graduatingStudentsSheet.getRange('C1:F1').setValue('Last updated: ' + Utilities.formatDate(new Date(), "GMT-7", "MM/dd HH:mm"));
//   }
// }