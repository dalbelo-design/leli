/**
 * LELI Contact Form — Google Apps Script Backend
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this entire file, replacing anything already there
 * 4. Save, then click Deploy > Manage Deployments
 *    - Edit your existing deployment (or create a new one)
 *    - Type: Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - Click Deploy and copy the URL into leli-contact.html
 */

const SHEET_NAME = 'Contacts';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create the sheet and header row if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        'Submitted At', 'First Name', 'Last Name', 'Email', 'Phone',
        'Organization', 'Reason for Contact', 'Message', 'How They Heard'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    // Parse URL-encoded form data (e.g. firstName=Maria&lastName=Gonzalez...)
    const data = e.parameter;

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

function sendNotification(data) {
  const recipient = 'loren@dalbert.design';
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

function doGet() {
  return ContentService
    .createTextOutput('LELI Contact Form endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
