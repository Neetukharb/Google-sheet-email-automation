function checkAllProjects() {
  const controlSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control");
  const data = controlSheet.getDataRange().getValues();

  const today = new Date();
  let currentMonth = today.getMonth(); 
  let checkMonth = currentMonth - 1;
  if (checkMonth < 0) checkMonth = 11;

  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];

  // UPDATED RANGE SEQUENCE (April-based structure)
  const monthRangeSequence = [
    "J5:K32",   // April
    "Q5:R32",   // May
    "X5:Y32",   // June
    "AH5:AI32", // July
    "AO5:AP32", // August
    "AV5:AW32", // September
    "BF5:BG32", // October
    "BM5:BN32", // November
    "BT5:BU32", // December
    "CD5:CE32", // January
    "CK5:CL32", // February
    "CR5:CS32"  // March
  ];

  for (let i = 1; i < data.length; i++) {
    let row = data[i];

    let sheetLink = row[2];
    let sheetName = row[3];
    let email = row[4];
    let startMonthName = row[5];
    let status = row[6];

    if (status !== "Active") continue;

    let startMonthIndex = monthNames.indexOf(startMonthName);

    // Key Logic: dynamic offset
    let monthOffset = (checkMonth - startMonthIndex + 12) % 12;

    let rangeToCheck = monthRangeSequence[monthOffset];

    try {
      let spreadsheet = SpreadsheetApp.openByUrl(sheetLink);
      let sheet = spreadsheet.getSheetByName(sheetName);

      let values = sheet.getRange(rangeToCheck).getValues();

      let isFilled = true;

      for (let r of values) {
        for (let c of r) {
          if (!c) {
            isFilled = false;
            break;
          }
        }
        if (!isFilled) break;
      }

      let emailStatus = "Not Sent";

      if (!isFilled) {
        sendReminder(email, monthNames[checkMonth], sheetLink);
        emailStatus = "Sent";
      }

      // Update Control Sheet
      controlSheet.getRange(i + 1, 8).setValue(new Date()); // Last Checked
      controlSheet.getRange(i + 1, 9).setValue(emailStatus);

    } catch (err) {
      Logger.log("Error at row " + i + ": " + err);
    }
  }
}


// Email Function
function sendReminder(email, monthName, sheetLink) {
  const subject = "Monthly Report Incomplete: Action Needed";

  const body = `
Dear Sir/Madam,

This is a gentle reminder that your monthly report for ${monthName} appears to be incomplete.

Please update the sheet at the earliest.

Link to the report:
${sheetLink}

Let me know if you need any support.

Warm regards,  
Neetu Kharb  
Database Management Executive
`;

  MailApp.sendEmail(email, subject, body);
}
