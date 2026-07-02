// ─── Google Apps Script ────────────────────────────────────
// 1. Go to sheets.google.com → create a new blank spreadsheet
// 2. Extensions → Apps Script
// 3. Paste this entire file (replace the default myFunction)
// 4. Click "Deploy" → New Deployment → Type: "Web app"
// 5. Set "Execute as" = Me, "Who has access" = Anyone
// 6. Copy the deployment URL — paste it into index.html under WAITLIST_URL
//
// Your sheet will auto-create columns: Timestamp | Email | Source
// Each signup also sends an email notification to NOTIFY_EMAIL

var NOTIFY_EMAIL = 'clintbailo94@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const email = data.email;

    if (!email || !email.includes('@')) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Invalid email' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Create header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.getRange('A1:C1').setValues([['Timestamp', 'Email', 'Source']]);
      sheet.getRange('A1:C1').setFontWeight('bold');
    }

    const source = data.source || 'landing-page';
    const timestamp = new Date().toISOString();

    sheet.appendRow([timestamp, email, source]);

    // Send email notification (wrapped so sheet logging still works if email fails)
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '[Tirado AI] New Waitlist Signup — ' + email,
        body: [
          'New waitlist signup on Tirado AI:',
          '',
          'Email:     ' + email,
          'Source:    ' + source,
          'Timestamp: ' + timestamp,
          '',
          '— Tirado AI Waitlist'
        ].join('\n')
      });
    } catch (mailErr) {
      console.error('Email send failed: ' + mailErr.message);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Tirado AI waitlist endpoint is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
