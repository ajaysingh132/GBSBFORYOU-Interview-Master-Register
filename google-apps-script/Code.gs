/**
 * GBSBFORYOU Interview Master Register — Google Apps Script endpoint
 * Deploy as Web app: Execute as Me; access Anyone (or your intended audience).
 * Store one row per question/answer in the spreadsheet linked to this script.
 */
function doPost(e) {
  try {
    var raw = e && e.parameter && e.parameter.payload;
    if (!raw) return output_({ ok: false, message: 'Missing payload' });
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Attach this script to the destination Google Sheet first.');
    var sheet = ss.getSheetByName('Interview Responses');
    if (!sheet) sheet = ss.insertSheet('Interview Responses');
    var headers = ['Submitted At', 'Interview ID', 'Title', 'Person', 'Interviewer', 'Topic', 'Respondent', 'Consent', 'Question No.', 'Question', 'Answer'];
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    var submitted = data.submittedAt || new Date().toISOString();
    (data.answers || []).forEach(function(a, i) {
      sheet.appendRow([submitted, data.interviewId || '', data.title || '', data.person || '', data.interviewer || '', data.topic || '', data.respondent || '', data.consent || '', i + 1, a.question || '', a.answer || '']);
    });
    return output_({ ok: true, message: 'Saved', interviewId: data.interviewId || '' });
  } catch (err) {
    return output_({ ok: false, message: String(err) });
  }
}
function doGet() {
  return ContentService.createTextOutput('GBSBFORYOU Interview endpoint is running. Submit interview data with POST.')
    .setMimeType(ContentService.MimeType.TEXT);
}
function output_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
