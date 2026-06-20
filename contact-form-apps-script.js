/**
 * LELI Contact Form — Google Apps Script Backend
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet (create one named "LELI Contact Form Submissions")
 * 2. Go to Extensions > Apps Script
 * 3. Paste this entire file, replacing anything already there
 * 4. Save, then click Deploy > New Deployment
 *    - Type: Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL and paste it into leli-contact.html
 *    replacing 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE'
 */

const SHEET_NAME = 'Contacts';  // Tab name in your Google Sheet

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create the sheet and header row if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Submitted At',
        'First Name',
        'Last Name',
        'Email',
        'Phone',
        'Organization',
        'Reason for Contact',
        'Message',
        'How They Heard',
      ]);
      // Bold the header row
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    }

    // Parse the incoming JSON
    const data = JSON.parse(e.postData.contents);

    // Append one row per submission
    sheet.appendRow([
      data.submittedAt    || new Date().toLocaleString(),
      data.firstName      || '',
      data.lastName       || '',
      data.email          || '',
      data.phone          || '',
      data.organization   || '',
      data.reason         || '',
      data.message        || '',
      data.hearAbout      || '',
    ]);

    // Optional: send a notification email
    sendNotification(data);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sends an email notification to the team when a new message arrives.
 * Remove or comment out this function if you don't want email alerts.
 */
function sendNotification(data) {
  const recipient = 'loren@dalbert.design';  // Change or add comma-separated addresses
  const subject   = `[LELI Contact] ${data.reason || 'New message'} from ${data.firstName} ${data.lastName}`;
  const body = `
New contact form submission on leli.design

Name:         ${data.firstName} ${data.lastName}
Email:        ${data.email}
Phone:        ${data.phone || '—'}
Organization: ${data.organization || '—'}
Reason:       ${data.reason}
Heard via:    ${data.hearAbout || '—'}
Submitted:    ${data.submittedAt}

--- Message ---
${data.message}
  `.trim();

  MailApp.sendEmail(recipient, subject, body);
}

/**
 * GET handler — returns a simple health-check so you can confirm the
 * deployment is live by visiting the URL in a browser.
 */
function doGet() {
  return ContentService
    .createTextOutput('LELI Contact Form endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
