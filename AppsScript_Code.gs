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
 * PAYMENT TRACKING: every new application also gets Fee Amount, Amount
 * Paid, Balance Remaining, Payment Status, Last Payment Date, and Receipt
 * Photo Link columns. Balance Remaining and Payment Status are formulas
 * that calculate themselves — the registrar only ever types into Fee
 * Amount, Amount Paid, Last Payment Date, and Receipt Photo Link.
 * Payment info is intentionally NOT shown on the public status.html page
 * — the reference-code + last-name lookup isn't a strong enough check to
 * safely expose financial data to anyone who guesses or overhears it.
 *
 * SETUP: this is a STANDALONE script (not bound to the Sheet via
 * Extensions -> Apps Script) — create it fresh at script.google.com/home
 * while logged into the account that should own it, paste this file in,
 * then Deploy as a Web App ("Anyone" access) -> copy the resulting URL
 * into register.html and status.html where marked
 * REPLACE-WITH-DEPLOYED-SCRIPT-ID. It connects to the Sheet by ID below,
 * so the deploying account just needs edit access to that Sheet — it
 * does not need to own it.
 *
 * DEPLOYMENT ACCOUNT: this runs under devt.liberia@gmail.com — DEVT's own
 * institutional account — so registration confirmation emails come
 * directly from DEVT, not from a Liberia Forward or personal address.
 * Confirmation emails display as "DEVT Registration System" via the name
 * option on MailApp.sendEmail below. (Earlier attempts under
 * ltotimeh@liberiaforward.org and mfoboi@liberiaforward.org appeared to
 * 404 on their public /exec URLs, but that turned out to be an artifact
 * of testing them from inside an automated browser session, not a real
 * problem with any account — always verify a freshly deployed Apps
 * Script URL from a normal, non-automated browser.)
 */

// The DEVT Registration Tracker Sheet's ID (from its URL). The account
// deploying this script needs at least edit access to this Sheet.
const SPREADSHEET_ID = "1U2GxNbM8Mcfn61iKEgy9PVYstC5oXhuoRhG7xDd5vPo";

const SHEET_NAME = "Sheet1";
const HEADERS = [
  "Reference Code", "Timestamp", "First Name", "Last Name",
  "Phone", "Email", "Program(s) Interested In", "Status", "Notes",
  "Enrolled in Classroom? (Y/N)",
  "Fee Amount", "Amount Paid", "Balance Remaining", "Payment Status",
  "Last Payment Date", "Receipt Photo Link"
];
// Column letters for the payment columns, used to build formulas below.
// K = Fee Amount, L = Amount Paid, M = Balance Remaining, N = Payment Status
const COL_FEE = "K", COL_PAID = "L", COL_BALANCE = "M", COL_PAYSTATUS = "N";

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const lastRow_ = sheet.getLastRow();
  const existing = lastRow_ < 2 ? [] : sheet.getRange(2, 1, lastRow_ - 1, 1)
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
    const newRow = sheet.getLastRow() + 1;

    // Balance Remaining and Payment Status are formulas, not typed-in values —
    // they recalculate automatically whenever the registrar edits Fee Amount
    // or Amount Paid. The registrar never has to touch these two columns.
    const balanceFormula =
      "=IF(" + COL_PAID + newRow + "=\"\",\"\"," +
      COL_FEE + newRow + "-" + COL_PAID + newRow + ")";
    const paymentStatusFormula =
      "=IF(" + COL_PAID + newRow + "=\"\",\"Unpaid\"," +
      "IF(" + COL_PAID + newRow + ">=" + COL_FEE + newRow + ",\"Paid in Full\",\"Partially Paid\"))";

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
      "N",
      "",                    // Fee Amount — registrar fills this in based on the program's cost
      "",                    // Amount Paid — registrar fills this in as payments come in
      balanceFormula,        // Balance Remaining — calculates itself
      paymentStatusFormula,  // Payment Status — calculates itself
      "",                    // Last Payment Date — registrar fills this in
      ""                     // Receipt Photo Link — paste the Drive link to the receipt photo
    ]);

    if (body.email) {
      try {
        MailApp.sendEmail({
          to: body.email,
          cc: "devt.liberia@gmail.com",
          replyTo: "devt.liberia@gmail.com",
          name: "DEVT Registration System",
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
