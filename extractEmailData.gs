function extractEmailData() {
    // Open the spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetData = sheet.getSheetByName('MailData');

    //Get Todays Date
    var todayDate = new Date();
    todayDate.setDate(todayDate.getDate() - 1); //Set date to 1 day before
    var formattedAfterDate = Utilities.formatDate(todayDate, Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Define your Gmail search query
    const searchQuery = 'from:example@mail.com subject:"Put Subject Here"' + ' after:' + formattedAfterDate;
    const threads = GmailApp.search(searchQuery);
    //const regexPattern = * Write your regex filter for email here *; 

    if (threads.length != 0) { //Check if emails detected
      threads.forEach(thread => {
          const messages = thread.getMessages();
          messages.forEach(message => {
              const date = message.getDate();
              const sender = message.getFrom();
              const subject = message.getSubject();
              const body = message.getPlainBody();

              //var match = body.match(regexPattern); //If using Regex pattern
              //var extractedInfo = match ? match[1] : "Not Found"; //If using Regex pattern

              // Append data to the spreadsheet
              //sheetData.appendRow([date, sender, subject, extractedInfo]); //If using Regex pattern
              sheetData.appendRow([date, sender, subject, body])
          });
      });
    } else {
      sheetData.appendRow(["No emails found today"]);   
    }

}
