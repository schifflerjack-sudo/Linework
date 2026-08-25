// Linework contact form handler
// Deploy as a Google Apps Script web app (Execute as: Me, Access: Anyone)

var SHEET_ID          = '1An--aHWbGxLW3q7XCeku7rbfjdVfrzG1OQGQSet-Omc';
var PROJECTS_FOLDER_ID = '1JYqF_-TAqCNEPdBsicDtAEsoCMT7sczf';
var NOTIFY_EMAIL      = 'jack@lineworksurveying.com';

// Cloudflare Turnstile: the secret key lives in Script Properties (Project
// Settings > Script Properties in the Apps Script editor), never here in
// source, since this file is committed to a repo.
var TURNSTILE_SECRET_PROPERTY = 'TURNSTILE_SECRET_KEY';

// True if the Turnstile token on the submission is valid. If the secret key
// hasn't been configured yet, this fails open (treats the submission as
// verified) and emails an alert, so a setup oversight doesn't silently
// drop every legitimate lead.
function isTurnstileVerified_(token) {
  var secret = PropertiesService.getScriptProperties().getProperty(TURNSTILE_SECRET_PROPERTY);
  if (!secret) {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: '[Linework] Turnstile not configured',
      body: 'The ' + TURNSTILE_SECRET_PROPERTY + ' script property is missing, so Turnstile ' +
            'verification is being skipped for all submissions. Set it under Project Settings > ' +
            'Script Properties in the Apps Script editor.'
    });
    return true;
  }
  if (!token) return false;

  try {
    var response = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      method: 'post',
      payload: { secret: secret, response: token },
      muteHttpExceptions: true
    });
    var result = JSON.parse(response.getContentText());
    return !!result.success;
  } catch (verifyErr) {
    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '[Linework] Turnstile verification error', body: verifyErr.toString() });
    return false;
  }
}

var SERVICE_LABELS = {
  'boundary':             'Boundary / Property Survey',
  'topographic':          'Topographic Survey',
  'construction-staking': 'Construction Staking',
  'alta':                 'ALTA / NSPS Survey',
  'lot-line':             'Lot Line Adjustment / Subdivision',
  'elevation-cert':       'Elevation Certificate',
  'easement':             'Easement Documentation',
  'condo':                'Condominium Conversion',
  'consultation':         'Consultation',
  'unsure':               'Not Sure'
};

var TIMELINE_LABELS = {
  'asap':    'As soon as possible',
  '1month':  'Within one month',
  '3months': 'Within three months',
  'flexible': 'Flexible'
};

function doPost(e) {
  try {
    var p = e.parameter;

    if (!isTurnstileVerified_(p['cf-turnstile-response'])) {
      // Failed or missing Turnstile token: pretend success so a bot can't
      // tell it was blocked, but skip the sheet write and emails.
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var firstName      = (p['first-name']        || '').trim();
    var lastName       = (p['last-name']         || '').trim();
    var clientName     = (firstName + ' ' + lastName).trim();
    var email          = (p['email']             || '').trim();
    var phone          = (p['phone']             || '').trim();
    var mailingAddress = (p['mailing-address']   || '').trim();
    var service        = SERVICE_LABELS[p['service']] || (p['service'] || '');
    var address        = (p['property-address']  || '').trim();
    var city           = (p['property-city']     || '').trim();
    var county         = (p['property-county']   || '').trim();
    var description    = (p['description']       || '').trim();
    var timeline       = TIMELINE_LABELS[p['timeline']] || (p['timeline'] || '');

    var propertyLocation = address + (city ? ', ' + city : '');

    var sheet   = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var lastRow = sheet.getLastRow();

    // Find the first row where column B (Date Received) is empty.
    // This matches pre-filled job numbers in column A to incoming submissions.
    // If no such row exists, append after the last row.
    var targetRow = lastRow + 1;
    if (lastRow > 1) {
      var colB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var i = 0; i < colB.length; i++) {
        if (!colB[i][0]) { targetRow = i + 2; break; }
      }
    }

    // Read the job number already in column A of the target row
    var jobNumber = sheet.getRange(targetRow, 1).getValue();

    // Write columns B-J only — column A (job number) is pre-filled manually
    // B: Date Received  C: Service Type  D: Timeline
    // E: Property Address + City  F: Property County
    // G: Client Name  H: Client Email  I: Client Mailing Address  J: Client Phone
    sheet.getRange(targetRow, 2, 1, 9).setValues([[
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      service,
      timeline,
      propertyLocation,
      county,
      clientName,
      email,
      mailingAddress,
      phone
    ]]);

    // Create project folder in Google Drive
    try {
      var folderName    = (jobNumber ? jobNumber + ' ' : '') + address;
      var projectFolder = DriveApp.getFolderById(PROJECTS_FOLDER_ID).createFolder(folderName);

      var mapsUrl = 'https://www.google.com/maps/search/?api=1&query=' +
                    encodeURIComponent(propertyLocation);
      projectFolder.createFile('Google Maps.url',
        '[InternetShortcut]\r\nURL=' + mapsUrl + '\r\n',
        MimeType.PLAIN_TEXT);

    } catch (driveErr) {
      MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '[Linework] Drive folder error', body: driveErr.toString() });
    }

    // Notification email to Linework
    var subject = 'New Inquiry: ' + clientName + ' (' + service + ')';
    var body = [
      'New project inquiry from the Linework website.',
      '',
      'CLIENT',
      'Name:             ' + clientName,
      'Email:            ' + email,
      'Phone:            ' + (phone || 'Not provided'),
      'Mailing Address:  ' + (mailingAddress || 'Not provided'),
      '',
      'PROPERTY',
      'Address: ' + propertyLocation,
      'County:  ' + (county || 'Not specified'),
      '',
      'PROJECT',
      'Service:  ' + service,
      'Timeline: ' + (timeline || 'Not specified'),
      '',
      'DESCRIPTION',
      description
    ].join('\n');

    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: body, replyTo: email });

    // Confirmation email to client
    var confirmBody = [
      'Hello ' + firstName + ',',
      '',
      'Thank you for reaching out to work with us at Linework Land Surveying. We have received your inquiry and will be in touch within five business days.',
      '',
      'In the meantime, if you have any questions you can reach us at:',
      '  Phone: (510) 224-4380',
      '  Email: jack@lineworksurveying.com',
      '',
      'We look forward to working with you.',
      '',
      'Linework Land Surveying',
      'Licensed Professional Land Surveyor',
      'Bay Area, California'
    ].join('\n');

    try {
      MailApp.sendEmail({ to: email, subject: 'We received your inquiry - Linework Land Surveying', body: confirmBody, replyTo: NOTIFY_EMAIL });
    } catch (clientErr) {
      MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '[Linework] Client email error', body: 'Client: ' + email + '\n\n' + clientErr.toString() });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function testDrive() {
  var folder = DriveApp.getFolderById(PROJECTS_FOLDER_ID);
  folder.createFolder('TEST - delete me');
  console.log('Drive OK');
}

function testEmail() {
  MailApp.sendEmail({ to: 'schifflerjack@gmail.com', subject: 'Test from Apps Script', body: 'Test email.' });
  console.log('Email OK');
}

// Run this once manually (select it in the function dropdown next to Run,
// then click Run) to grant the script.external_request permission that
// UrlFetchApp needs for Turnstile verification. Without this, doPost fails
// with "You do not have permission to call UrlFetchApp.fetch" the first
// time it tries to verify a token. Approving the consent prompt here
// authorizes the whole script project, so no redeploy is needed afterward.
function testTurnstile() {
  var response = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
    method: 'post',
    payload: { secret: 'test', response: 'test' },
    muteHttpExceptions: true
  });
  console.log(response.getContentText());
}
