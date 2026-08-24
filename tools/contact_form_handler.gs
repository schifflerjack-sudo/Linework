// Linework contact form handler
// Deploy as a Google Apps Script web app (Execute as: Me, Access: Anyone)

var SHEET_ID          = '1An--aHWbGxLW3q7XCeku7rbfjdVfrzG1OQGQSet-Omc';
var PROJECTS_FOLDER_ID = '1JYqF_-TAqCNEPdBsicDtAEsoCMT7sczf';
var NOTIFY_EMAIL      = 'jack@lineworksurveying.com';

// --- Spam filtering ---
// Submissions faster than this are almost certainly scripted bots, not a
// person filling out a multi-field form.
var MIN_FILL_SECONDS  = 3;
// Flag (don't block) submissions whose description looks spammy, so real
// leads that happen to include a link aren't silently dropped.
var SPAM_URL_THRESHOLD = 2;
var SPAM_KEYWORDS = [
  'seo services', 'backlink', 'link building', 'crypto', 'bitcoin',
  'forex', 'casino', 'viagra', 'cialis', 'weight loss', 'work from home',
  'guest post', 'increase your ranking', 'social media marketing'
];
// Counties actually served - see the contact page's "Service Area" list.
var SERVICE_COUNTIES = [
  'alameda', 'contra costa', 'marin', 'san mateo', 'san francisco', 'santa clara'
];

// True if the honeypot field was filled in (real users never see or fill it).
function isHoneypotTriggered_(p) {
  return !!(p['bot-field'] || '').trim();
}

// True if the form was submitted implausibly fast after it loaded.
function isSubmittedTooFast_(p) {
  var ts = Number(p['form-timestamp']);
  if (!ts) return false; // missing timestamp (e.g. JS disabled) - don't block
  var elapsedSeconds = (Date.now() - ts) / 1000;
  return elapsedSeconds >= 0 && elapsedSeconds < MIN_FILL_SECONDS;
}

// True if the free-text description looks like spam content.
function looksLikeSpamContent_(text) {
  var lower = (text || '').toLowerCase();
  var urlMatches = lower.match(/https?:\/\/|www\./g) || [];
  var keywordHit = SPAM_KEYWORDS.some(function (k) { return lower.indexOf(k) !== -1; });
  return urlMatches.length >= SPAM_URL_THRESHOLD || keywordHit;
}

// True if the county looks like one of the Bay Area counties actually served.
function isRecognizedServiceCounty_(county) {
  var normalized = (county || '').toLowerCase().replace(/\s*county\s*$/, '').trim();
  return SERVICE_COUNTIES.indexOf(normalized) !== -1;
}

// True if the phone number looks like a US number (10 digits, or 11 starting
// with 1). Blank is treated as fine since the field is optional. A trailing
// extension (" ext 2", " x2", " #2") is stripped first so it isn't counted
// as part of the number.
function looksLikeUsPhone_(phone) {
  var withoutExtension = (phone || '').replace(/\b(ext\.?|x|#)\s*\d+\s*$/i, '');
  var digits = withoutExtension.replace(/\D/g, '');
  if (!digits) return true;
  return digits.length === 10 || (digits.length === 11 && digits.charAt(0) === '1');
}

// True if a submission looks like the scripted-bot spam pattern seen in
// practice: an identical first/last name (e.g. "Robertchats Robertchats"),
// plus at least one of a property county outside the service area or a
// non-US phone number. The identical-name signal is required on its own
// terms, not just scored alongside the others: a county outside the service
// area or a non-US phone each have common innocent explanations (a
// referral-seeking lead from outside the area - which the contact page
// explicitly invites - or a client traveling abroad), so either alone, or
// even both together, is left alone. Only combined with a duplicated name
// does it cross into "reject" territory.
function isLikelySpamSubmission_(p, county) {
  var first = (p['first-name'] || '').trim().toLowerCase();
  var last  = (p['last-name']  || '').trim().toLowerCase();
  var nameRepeated = !!(first && first === last);
  if (!nameRepeated) return false;

  var otherSignals = 0;
  if (!isRecognizedServiceCounty_(county)) otherSignals++;
  if (!looksLikeUsPhone_(p['phone'])) otherSignals++;
  return otherSignals >= 1;
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

    if (isHoneypotTriggered_(p) || isSubmittedTooFast_(p)) {
      // Likely a bot: silently pretend success so it doesn't retry or
      // adjust, but skip the sheet write, folder creation, and emails.
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

    if (isLikelySpamSubmission_(p, county)) {
      // Matches the observed bot pattern: treat as spam the same way as the
      // honeypot/timing checks above.
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

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
    var spamFlag = looksLikeSpamContent_(description) ? '[Possible Spam] ' : '';
    var subject = spamFlag + 'New Inquiry: ' + clientName + ' (' + service + ')';
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
