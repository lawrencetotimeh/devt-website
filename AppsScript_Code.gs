/**
 * DEVT Registration System — Google Apps Script
 * ------------------------------------------------
 * This script is the small "engine" that connects the website's
 * registration form and status-check page to the DEVT Registration
 * Tracker Google Sheet. It does two things:
 *
 *   1. doPost()  — runs when a student submits the registration form
 *                  on register.html. Writes a new row to the Sheet,
 *                  generates a reference code, emails it if an email
 *                  was given, and returns the code to the website.
 *
 *   2. doGet()   — runs when a student checks their status on
 *                  status.html. Looks up their reference code + last
 *                  name in the Sheet and returns their current status.
 *
 * SETUP: see the Registrar & Mohamed Guide, Step 4, for exactly how to
 * install and deploy this. In short: open the DEVT Registration Tracker
 * Sheet -> Extensions -> Apps Script -> paste this file in -> Deploy as
 * a Web App ("Anyone" access) -> copy the resulting URL into
 * register.html and status.html where marked REPLACE-WITH-DEPLOYED-SCRIPT-ID.
 */

const SHEET_NAME = "Sheet1";
const HEADERS = [
  "Reference Code", "Timestamp", "First Name", "Last Name",
  "Phone", "Email", "Program(s) Interested In", "Status", "Notes",
  "Enrolled in Classroom? (Y/N)"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  // Make sure the header row exists and is correct.
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow.join("") === "") {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function generateReferenceCode_(sheet) {
  // Simple, short, not easily guessable: DEVT- + 4 random digits,
  // re-rolled if it happens to collide with an existing code.
  const existing = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1)
    .getValues().flat().map(String);
  let code;
  do {
    code = "DEVT-" + Math.floor(1000 + Math.random() * 9000);
  } while (existing.indexOf(code) !== -1);
  return code;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== "register") {
      return jsonResponse_({ success: false, error: "Unknown action" });
    }
    if (!body.firstName || !body.lastName || !body.phone || !body.program) {
      return jsonResponse_({ success: false, error: "Missing required fields" });
    }

    const sheet = getSheet_();
    const code = generateReferenceCode_(sheet);
    const timestamp = new Date();

    sheet.appendRow([
      code,
      timestamp,
      body.firstName,
      body.lastName,
      body.phone,
      body.email || "",
      body.program,
      "Received",
      body.notes || "",
      "N"
    ]);

    if (body.email) {
      try {
        MailApp.sendEmail({
          to: body.email,
          subject: "Your DEVT application — reference code " + code,
          body:
            "Hi " + body.firstName + ",\n\n" +
            "Thank you for applying to DEVT (Ducor Institute of Vocational and Technical Development).\n\n" +
            "Your reference code is: " + code + "\n\n" +
            "You can check your application status anytime on our website's \"Check Your Status\" page using this code and your last name.\n\n" +
            "— DEVT"
        });
      } catch (mailErr) {
        // Don't fail the whole registration if the email send fails —
        // the student still has the code on-screen.
      }
    }

    return jsonResponse_({ success: true, referenceCode: code });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

// Friendly, plain-language text shown alongside each status.
const STATUS_MESSAGES = {
  "Received": "We've received your application. Our registrar will review it soon.",
  "Under Review": "Our registrar is currently reviewing your application.",
  "Accepted": "Congratulations — you've been accepted! We'll be in touch about next steps.",
  "Enrolled": "You're enrolled. Check with the registrar's office for your Google Classroom access.",
  "Waitlisted": "You're on our waitlist for this program. We'll reach out if a spot opens.",
  "Not Accepted": "Thank you for applying. Unfortunately we're unable to offer you a spot at this time."
};

function doGet(e) {
  try {
    const code = (e.parameter.code || "").trim().toUpperCase();
    const lastName = (e.parameter.lastName || "").trim().toLowerCase();
    if (!code || !lastName) {
      return jsonResponse_({ found: false });
    }

    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return jsonResponse_({ found: false });

    const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (const row of rows) {
      const rowCode = String(row[0]).trim().toUpperCase();
      const rowLastName = String(row[3]).trim().toLowerCase();
      if (rowCode === code && rowLastName === lastName) {
        const status = row[7] || "Received";
        return jsonResponse_({
          found: true,
          status: status,
          message: STATUS_MESSAGES[status] || ""
        });
      }
    }
    return jsonResponse_({ found: false });
  } catch (err) {
    return jsonResponse_({ found: false, error: String(err) });
  }
}
