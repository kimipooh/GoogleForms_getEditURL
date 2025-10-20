/**
 * Fetches the edit URL for all existing form responses and populates them
 * into Column ?? (specified by "saveTocolumn" of the active spreadsheet.
 * This function must be run manually from the script editor.
 
 Kimiya Kitani
 
 20 October 2025: Version 1.0
 */
// A = 1, B = 2, .... D = 4, .... 
// Please specify the column number for which you want to put an edit URL. Please note that the selected column value is overwritten.
const saveTocolumn = 0;

function getPastEditUrls() {
  // ▼▼▼ CONFIGURATION ▼▼▼
  // Please paste your Google Form's ID between the single quotes.
  // You can find the ID in the form's URL (between /d/ and /edit).
  const formId = 'YOUR_FORM_ID_HERE'; 
  // ▲▲▲ END CONFIGURATION ▲▲▲
  
  try {
    // Open the Google Form by its ID.
    const form = FormApp.openById(formId);
    if (!form) {
      throw new Error('Could not find the form with the specified ID.');
    }
    if (saveTocolumn <= 0){
      throw new Error('Please specify "saveTocolumn" elements in Google Apps Script.')   
    }
    // Get the active sheet in the currently open spreadsheet.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Retrieve all responses from the form.
    const formResponses = form.getResponses();
    // Retrieve all data from the sheet, excluding the header row.
    const sheetData = sheet.getDataRange().getValues();
    sheetData.shift(); // Remove the header row from the data array.
    
    // Create a map to quickly look up an edit URL by its timestamp.
    // This is much more efficient than searching all responses for each sheet row.
    const responseMap = {};
    formResponses.forEach(response => {
      // Format the timestamp to match the spreadsheet's format for reliable matching.
      const timestamp = Utilities.formatDate(response.getTimestamp(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
      responseMap[timestamp] = response.getEditResponseUrl();
    });
    
    // Iterate over each row of data in the spreadsheet.
    sheetData.forEach((row, index) => {
      // Get the timestamp from the first column (Column A).
      const sheetTimestamp = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
      
      // If a matching response is found in our map, get the URL.
      if (responseMap[sheetTimestamp]) {
        // Calculate the correct row number (index is 0-based, plus header row, so +2).
        const targetRow = index + 2;
        // Set the edit URL value in the 4th column (Column D).
        sheet.getRange(targetRow, saveTocolumn).setValue(responseMap[sheetTimestamp]);
        console.log(targetRow+": done.");
      }
    });
    
    // Show a confirmation message to the user.
    //SpreadsheetApp.getUi().alert('Processing complete for existing entries.');
    
  } catch (error) {
    // Show an error message if something went wrong.
    SpreadsheetApp.getUi().alert('An error occurred: ' + error.toString());
    console.error('Failed to batch process URLs: ' + error.toString());
  }
}
