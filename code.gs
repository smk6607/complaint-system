// Smart Grievance Portal - PDA College of Engineering
// Google Apps Script backend (bound to the Google Sheet that stores complaints)

var STATUS_COL = 10; // column J = Status
var COLLEGE = "PDA College of Engineering";

// ---------- NEW COMPLAINT ----------
function doPost(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Error: No data received.");
  }
  var data = e.parameter;

  // Honeypot: bots fill this hidden field, humans don't
  if (data.website) return ContentService.createTextOutput("ok");

  if (!data.category || !data.details || !data.department || !data.email) {
    return ContentService.createTextOutput("Error: Missing required fields.");
  }

  var email = String(data.email).trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return ContentService.createTextOutput("Error: Invalid email address.");
  }

  var priority = ["Low", "Medium", "High"].indexOf(data.priority) >= 0 ? data.priority : "Low";
  var affected = Math.min(100000, Math.max(1, parseInt(data.affected, 10) || 1));

  var trackingId;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    trackingId = pickTrackingId_(sheet, data.trackingId);

    sheet.appendRow([
      new Date(),                    // A: Timestamp
      clean_(data.category, 100),    // B: Category
      clean_(data.details, 5000),    // C: Details
      clean_(data.location, 200),    // D: Location
      priority,                      // E: Priority
      clean_(data.department, 100),  // F: Department
      email,                         // G: Email
      affected,                      // H: Affected Count
      trackingId,                    // I: Tracking ID
      "Pending"                      // J: Status
    ]);
  } finally {
    lock.releaseLock();
  }

  try {
    var body =
      "Hello,\n\n" +
      "Your complaint regarding '" + clean_(data.category, 100) + "' has been registered and sent to the " +
      clean_(data.department, 100) + " department.\n\n" +
      "YOUR TRACKING ID: " + trackingId + "\n\n" +
      "Please keep this ID and follow it on the Live Tracker for further updates. " +
      "We will also email you whenever the status changes.\n\n" +
      "Thank you,\nSmart Grievance Portal\n" + COLLEGE;
    MailApp.sendEmail(email, "Complaint Registered - " + trackingId + " | PDA College", body);
  } catch (err) {
    console.error("Email failed: " + err);
  }

  return ContentService.createTextOutput("Success");
}

// ---------- STATUS CHANGE EMAIL ----------
// Runs when the Status column is edited (needs an installable trigger: run setupTrigger once)
function onStatusEdit(e) {
  if (!e || !e.range) return;
  var row = e.range.getRow();
  if (row < 2 || e.range.getColumn() !== STATUS_COL) return;

  var r = e.range.getSheet().getRange(row, 1, 1, 10).getValues()[0];
  var category = r[1], email = r[6], trackingId = r[8];
  var status = e.value || r[9];
  if (!email) return;

  var body =
    "Hello,\n\nThere is an update on your complaint regarding '" + category + "'.\n\n" +
    "Tracking ID: " + trackingId + "\n" +
    "New status: " + status + "\n\n" +
    "You can see full details on the Live Tracker.\n\n" +
    "Smart Grievance Portal\n" + COLLEGE;
  try {
    MailApp.sendEmail(email, "Complaint Update (" + status + ") - " + trackingId, body);
  } catch (err) {
    console.error("Email failed: " + err);
  }
}

// ---------- SETUP HELPERS (run once from the editor) ----------

// Creates the installable edit trigger for onStatusEdit (skips if it already exists)
function setupTrigger() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === "onStatusEdit";
  });
  if (!exists) {
    ScriptApp.newTrigger("onStatusEdit")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  }
}

// Grants Gmail permission by sending a test mail to yourself
function testEmail() {
  MailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    "Smart Grievance Portal - test email",
    "Email permission is working."
  );
}

// ---------- INTERNAL HELPERS ----------

// Trim, cap length, and neutralise spreadsheet formula injection (=, +, -, @)
function clean_(value, maxLen) {
  var s = String(value || "").trim().substring(0, maxLen);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Use the client's tracking ID only if it is well-formed and unused; otherwise generate one
function pickTrackingId_(sheet, requested) {
  var existing = {};
  var last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 9, last - 1, 1).getValues().forEach(function (row) {
      existing[String(row[0])] = true;
    });
  }
  var id = String(requested || "").trim();
  if (/^[A-Za-z0-9-]{4,30}$/.test(id) && !existing[id]) return id;

  do {
    id = "PDA-" + Utilities.formatDate(new Date(), "Asia/Kolkata", "yyMMdd") + "-" +
         Math.floor(1000 + Math.random() * 9000);
  } while (existing[id]);
  return id;
}
