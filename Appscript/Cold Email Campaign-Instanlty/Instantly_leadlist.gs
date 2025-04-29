/**
 * Filters and appends unique emails to a "Filtered Emails" sheet based on event counts.
 */
function filterAndAppendUniqueEmails() {
  // Get the active spreadsheet and the "Sheet1" sheet.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  // Check if "Sheet1" exists.
  if (!sheet) {
    // Log an error message if "Sheet1" is not found.
    Logger.log("Sheet1 not found.");
    return;
  }
  // Get all data from the sheet.
  var data = sheet.getDataRange().getValues();

  // Define the columns to copy to the "Filtered Emails" sheet.
  var columnsToCopy = ["LastName", "FirstName", "EmailAddress"];

  // Find the index of each column to copy.
  var headers = data[0];
  var columnIndexes = columnsToCopy.map(col => headers.indexOf(col));

  // Check if all required columns are found.
  if (columnIndexes.includes(-1)) {
    // Log an error message if one or more required columns are not found.
    Logger.log("One or more required columns not found.");
    return;
  }

  // Get the index of the "EmailAddress" and "Event Type" columns.
  var emailIndex = headers.indexOf("EmailAddress");
  var eventIndex = headers.indexOf("Event Type");

  // Create objects to track email event counts and store unique email data.
  var emailCounts = {}; // Track event counts for each email.
  var emailData = {}; // Store unique email data.

  // Iterate over the data to count occurrences of "email_opened" and "link_clicked" events.
  for (var i = 1; i < data.length; i++) {
    // Get the email address and event type for the current row.
    var email = data[i][emailIndex];
    var eventType = data[i][eventIndex];

    // Initialize the email counts object if it doesn't exist.
    if (!emailCounts[email]) {
      emailCounts[email] = { opened: 0, clicked: 0 };
    }

    // Increment the "opened" count if the event type is "email_opened".
    if (eventType === "email_opened") {
      emailCounts[email].opened += 1;
      // Store the selected columns for the email if not already stored.
      if (!emailData[email]) {
        emailData[email] = columnIndexes.map(idx => data[i][idx]); // Store selected columns
      }
    }

    // Increment the "clicked" count if the event type is "link_clicked".
    if (eventType === "link_clicked") {
      emailCounts[email].clicked += 1;
      // Store the selected columns for the email if not already stored.
      if (!emailData[email]) {
        emailData[email] = columnIndexes.map(idx => data[i][idx]); // Store selected columns
      }
    }
  }

  // Open or create the "Filtered Emails" sheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = ss.getSheetByName("Filtered Emails");
  // If the sheet doesn't exist, create it.
  if (!newSheet) {
    newSheet = ss.insertSheet("Filtered Emails");
    newSheet.appendRow(columnsToCopy); // Add headers only if creating a new sheet
  }

  // Get existing emails in the "Filtered Emails" sheet.
  var existingEmails = newSheet.getDataRange().getValues().map(row => row[2]); // Use EmailAddress as a unique identifier

  // Filter for new unique emails to add to the "Filtered Emails" sheet.
  var newRows = [];
  for (var email in emailCounts) {
    // Check if the email has been opened more than once or clicked at least once, and if it's not already in the "Filtered Emails" sheet.
    if ((emailCounts[email].opened > 1 || emailCounts[email].clicked >= 1) && !existingEmails.includes(email)) {
      newRows.push(emailData[email]); // Add only new unique emails with selected columns
    }
  }

  // Append the new unique rows to the "Filtered Emails" sheet.
  if (newRows.length > 0) {
    newSheet.getRange(newSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    // Log the number of new unique emails added.
    Logger.log(newRows.length + " new unique emails added.");
  } else {
    // Log a message if no new unique emails were added.
    Logger.log("No new unique emails to add.");
  }
}
