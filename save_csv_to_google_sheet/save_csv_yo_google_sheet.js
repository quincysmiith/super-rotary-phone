function importCSVFromGmail() {
  // To help out with compiling figures for the flash report each Monday
  // Function to load the latest csv email attachment from a Domo report.
  // Domo card that is sent to email: https://citrusad-cartology.domo.com/page/339676999/kpis/details/1398047832
  // Data is scheduled to be sent everyday
  // To see schedule: https://citrusad-cartology.domo.com/scheduled-reports
  // Data in attached CSV file is saved to the Google sheet: https://docs.google.com/spreadsheets/d/1ul0xk36R73TsjPHagbTs2nJ54i7An41b7OZq_YevMcw/edit#gid=0


  var threads = GmailApp.search("from:notifications@domo.com subject:\"Report - Previous 7 days Ad revenue (excl today)\"");

  // get latest message in thread
  var message = threads[0].getMessages()[0];
  var sentDate = message.getDate()

  // Loop through all message attachments
  var i;
  for (i = 0; i < message.getAttachments().length; i++) {
    var attachment = message.getAttachments()[i];

    // Is the attachment a CSV file
    // There should only be 1 csv attached
    // Behaviour of script may be unexpected if multiple csv's are attached.
    if (attachment.getContentType() === "text/csv") {

      // Select Google sheet and first tab
      var spreadsheet = SpreadsheetApp.openById("1ul0xk36R73TsjPHagbTs2nJ54i7An41b7OZq_YevMcw");
      var sheet = SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);

      var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");

      // Remember to clear the content of the sheet before importing new data
      sheet.clearContents().clearFormats();
      sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

      console.log("Data saved to spreadsheet from report sent on " + sentDate)


      // Enter value of last run date time in spreadsheet
      my_range = sheet.getRange('A13');
      let currentDate = new Date();
      let cDay = currentDate.getDate();
      my_range.setValue("Data last updated at: " + currentDate);


      // Enter the details of when data was sent from Domo into spreadsheet
      my_range = sheet.getRange('A14');
      my_range.setValue("Data saved to spreadsheet from report sent on " + sentDate);

    }
  }
}