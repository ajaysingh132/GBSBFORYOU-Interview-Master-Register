/**
 * GBSBFORYOU Interview + evidence upload endpoint.
 * Attach this script to the destination Google Sheet.
 * Deploy as Web app (Execute as: Me). Anyone access allows public submissions;
 * keep the URL private from spam and do not make uploaded files public automatically.
 */
function doPost(e) {
  try {
    var raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('Missing payload');
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Attach this script to the destination Google Sheet first.');

    var folder = getOrCreateFolder_('GBSBFORYOU Interview Evidence');
    var master = getSheet_(ss, 'Interview Master', [
      'Submitted At','Interview ID','Title','Person','Interviewer','Topic','Respondent','Consent','Answer Count'
    ]);
    var responses = getSheet_(ss, 'Interview Responses', [
      'Submitted At','Interview ID','Project ID','Respondent','Consent','Question No.','Question ID','Question','Answer','Attachment Count','Asset IDs','Drive File Links','Attachment Notes'
    ]);
    var assets = getSheet_(ss, 'Asset Register', [
      'Asset ID','Submitted At','Project ID','Interview ID','Question No.','Question ID','Original File Name','MIME Type','File Size Bytes','Drive File ID','Drive File URL','Description','Uploaded By'
    ]);

    var submitted = data.submittedAt || new Date().toISOString();
    var interviewId = data.interviewId || ('INT-' + new Date().getTime());
    var projectId = data.projectId || 'GBS-BOOK-HR001';
    var totalAnswers = (data.answers || []).length;
    master.appendRow([submitted, interviewId, data.title || '', data.person || '', data.interviewer || '', data.topic || '', data.respondent || '', data.consent || '', totalAnswers]);

    (data.answers || []).forEach(function(a, i) {
      var qNo = a.questionNo || (i + 1);
      var qId = a.questionId || ('Q-' + ('000' + qNo).slice(-3));
      var fileLinks = [];
      var assetIds = [];
      var notes = [];
      (a.files || []).forEach(function(file, fileIndex) {
        if (!file || !file.dataBase64 || !file.name) return;
        var bytes = Utilities.base64Decode(file.dataBase64);
        var blob = Utilities.newBlob(bytes, file.type || 'application/octet-stream', sanitizeName_(file.name));
        var driveFile = folder.createFile(blob);
        var assetId = 'AST-' + slug_(interviewId) + '-Q' + ('000' + qNo).slice(-3) + '-' + new Date().getTime().toString().slice(-6) + ('0' + (fileIndex + 1)).slice(-2);
        var url = driveFile.getUrl();
        fileLinks.push(url);
        assetIds.push(assetId);
        assets.appendRow([assetId, submitted, projectId, interviewId, qNo, qId, file.name, file.type || '', file.size || bytes.length, driveFile.getId(), url, a.attachmentNote || '', data.respondent || '']);
      });
      notes.push(a.attachmentNote || '');
      responses.appendRow([
        submitted, interviewId, projectId, data.respondent || '', data.consent || '',
        qNo, qId, a.question || '', a.answer || '', fileLinks.length,
        assetIds.join('\n'), fileLinks.join('\n'), notes.join('\n')
      ]);
    });
    return output_({ok:true, message:'Saved', interviewId:interviewId, answerCount:totalAnswers, filesSaved:countFiles_(data.answers || [])});
  } catch (err) {
    return output_({ok:false, message:String(err)});
  }
}
function doGet() {
  return ContentService.createTextOutput('GBSBFORYOU interview endpoint is running. Use the shareable questionnaire page and submit by POST.')
    .setMimeType(ContentService.MimeType.TEXT);
}
function getSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}
function getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function sanitizeName_(name) {
  return String(name || 'upload').replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 150);
}
function slug_(s) {
  return String(s || 'INT').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(-12) || 'INT';
}
function countFiles_(answers) {
  return answers.reduce(function(n, a) { return n + ((a.files || []).filter(function(f){return f && f.dataBase64;}).length); }, 0);
}
function output_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
