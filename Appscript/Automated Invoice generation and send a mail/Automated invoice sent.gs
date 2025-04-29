function sendInvoice() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = sheet.getSheetByName("Invoice"); // Ensure this matches your sheet name
  const clientDetailsSheet = sheet.getSheetByName("ClientDetails"); // Ensure this sheet contains client data

  if (!templateSheet || !clientDetailsSheet) {
    throw new Error("Ensure both 'Invoice' and 'ClientDetails' sheets exist.");
  }

  const clients = clientDetailsSheet.getDataRange().getValues(); // Fetch all client data
  const today = new Date();
  const invoiceDate = formatDate(today); // Current date in 'DD/MM/YYYY' format

  // Fetch the stored invoice tracker (persistent across runs)
  const properties = PropertiesService.getScriptProperties();
  let invoiceTracker = JSON.parse(properties.getProperty("invoiceTracker") || "{}");

  const emailRecipients = []; // Array to hold all recipients

  // Loop through each client (start from index 1 to skip header row)
  for (let i = 1; i < clients.length; i++) {
    const [name, address, emails, project, amount] = clients[i]; // Extract client details
    if (!name || !address || !emails || !project || !amount) {
      Logger.log(`Skipping row ${i + 1} due to missing data.`);
      continue; // Skip rows with missing data
    }

    const invoiceNumber = generateInvoiceNumber(name, invoiceTracker);

    // Generate invoice PDF for the current client
    const pdf = createInvoicePDF(templateSheet, name, address, project, invoiceDate, invoiceNumber, amount);

    // Send email to all listed addresses
    const emailList = emails.split(',').map(email => email.trim()); // Split multiple emails by comma
    emailRecipients.push(...emailList); // Add to the list of recipients

    // Remove duplicate emails from the list
    const uniqueEmailRecipients = [...new Set(emailRecipients)];

      if (pdf) {
        GmailApp.sendEmail(uniqueEmailRecipients, `${project} Invoice: ${invoiceNumber}`, 
          `Hello,\n\nPlease find attached invoice for ${project}.\n\nBest regards,\nBeena Thakkar`,
          {
            attachments: [pdf],
            name: "Beena Thakkar"
          });
        Logger.log(`Invoice sent to ${uniqueEmailRecipients} for ${project}.`);
      } else {
        Logger.log(`Failed to create invoice for ${name}.`);
      }
  }

  // Save the updated invoice tracker back to properties
  properties.setProperty("invoiceTracker", JSON.stringify(invoiceTracker));
}

// Helper to format the date
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// Helper to generate invoice number
function generateInvoiceNumber(clientName, tracker) {
  if (!tracker[clientName]) {
    tracker[clientName] = 1; // Start from 1 if no entry exists for the client
  }
  const prefix = clientName.charAt(0).toUpperCase(); // Use the first letter of the client name
  const invoiceNumber = `${prefix}${tracker[clientName].toString().padStart(3, "0")}`; // Format as "M001"
  tracker[clientName]++; // Increment the number for the next invoice
  return invoiceNumber;
}

// Updated `createInvoicePDF` to include the address
function createInvoicePDF(templateSheet, name, address, project, invoiceDate, invoiceNumber, amount) {
  const tempSheet = templateSheet.copyTo(templateSheet.getParent()); // Copy the template sheet
  tempSheet.setName(`Invoice_${invoiceNumber}`); // Rename temporary sheet

  // Update dynamic fields in the invoice
  tempSheet.getRange("C9").setValue(invoiceDate); // Update {{invoiceDate}}
  tempSheet.getRange("F12").setValue(invoiceNumber); // Update {{invoiceNumber}}
  tempSheet.getRange("B12").setValue(name); // Update Client Name
  tempSheet.getRange("B13:C13").setValue(address); // Update Address
  tempSheet.getRange("D15:G15").setValue(project); // Update Project
  tempSheet.getRange("B19:D19").setValue(project);
  tempSheet.getRange("G19").setValue(amount); // Update Total Price

  // Force changes to be saved before generating the PDF
  SpreadsheetApp.flush();

  // Get the spreadsheet ID and the sheet ID
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetId = tempSheet.getSheetId();

  // Construct the export URL for the temporary sheet
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${sheetId}&size=A4&portrait=true&fitw=true&top_margin=0.75&bottom_margin=0.75&left_margin=0.75&right_margin=0.75&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

  // Fetch the PDF content
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  const pdfBlob = response.getBlob().setName(`Invoice_${invoiceNumber}.pdf`);

  // Delete the temporary sheet after generating the PDF
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);

  // Return the PDF blob for immediate use (e.g., email attachment)
  return pdfBlob;
}
