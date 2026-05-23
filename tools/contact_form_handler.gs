// Linework contact form handler
// Deploy as a Google Apps Script web app (Execute as: Me, Access: Anyone)

var SHEET_ID = '1An--aHWbGxLW3q7XCeku7rbfjdVfrzG1OQGQSet-Omc';
var NOTIFY_EMAIL = 'jack@lineworksurveying.com';

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

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    // Column order:
    // A: Project Number (fill in manually)  B: Date Received  C: Service Type  D: Timeline
    // E: Property Address + City  F: Property County
    // G: Client Name  H: Client Email  I: Client Mailing Address  J: Client Phone
    // K-Q: status columns (Sent to Client, Needs Filing, Submitted to County,
    //       Date Submitted, Date Redlines Received, Filing Status, Recording Info)
    sheet.appendRow([
      '',
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      service,
      timeline,
      propertyLocation,
      county,
      clientName,
      email,
      mailingAddress,
      phone,
      '', '', '', '', '', '', ''
    ]);

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
      description,
      '',
      '---',
      'Job log: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
    ].join('\n');

    try {
      MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: body, replyTo: email });
    } catch (adminErr) {}

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
    } catch (clientErr) {}

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
