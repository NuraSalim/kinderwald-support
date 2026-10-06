const SHEET_NAME = "Tickets";
const NOTIFY_EMAIL = "YOUR_EMAIL@example.com";

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Ticket","Created At","Parent","Child","Mobile","Email",
      "Issue Type","Priority","Description","Status"
    ]);
  }
}

function doPost(e) {
  try {
    setup();
    const data = JSON.parse(e.postData.contents);
    const ticket = "KW-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
    const now = new Date();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    sheet.appendRow([
      ticket, now, data.parentName || "", data.childName || "", data.mobile || "",
      data.email || "", data.issueType || "", data.priority || "Normal",
      data.description || "", "Open"
    ]);

    const subject = "Kinderwald Nursery - New Support Ticket " + ticket;
    const body =
      "New technical support request\n\n" +
      "Ticket: " + ticket + "\n" +
      "Parent: " + (data.parentName || "") + "\n" +
      "Child: " + (data.childName || "") + "\n" +
      "Mobile: " + (data.mobile || "") + "\n" +
      "Email: " + (data.email || "") + "\n" +
      "Issue: " + (data.issueType || "") + "\n" +
      "Priority: " + (data.priority || "") + "\n\n" +
      "Description:\n" + (data.description || "");

    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);

    return json({ok:true, ticket:ticket});
  } catch (err) {
    return json({ok:false, error:String(err)});
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
