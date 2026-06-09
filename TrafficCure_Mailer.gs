/*** ============ CONFIG, set the sending speed you want ============ ***/
const BATCH_SIZE          = 5;    // emails per batch
const BATCH_INTERVAL_MIN  = 1;    // minutes between batches
const SHEET_NAME          = 'Sheet1';

// Column headers expected in row 1 of the sheet:
// Name | Email | Company | Subject | Body | Status | DraftId | SentAt
/*** =============================================================== ***/


/** STEP A, run this first. Creates a Gmail DRAFT for every contact. Nothing sends yet. */
function createDrafts() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const head  = data[0];
  const col   = name => head.indexOf(name);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col('Status')] === 'DRAFT' || row[col('Status')] === 'SENT') continue; // skip done
    const to = row[col('Email')];
    if (!to) continue;

    const subject = row[col('Subject')];
    const body    = row[col('Body')];
    if (!subject || !body) continue;

    const htmlBody = body.replace(/\n/g, '<br>');
    const draft = GmailApp.createDraft(to, subject, body, { htmlBody: htmlBody });
    sheet.getRange(i + 1, col('Status') + 1).setValue('DRAFT');
    sheet.getRange(i + 1, col('DraftId') + 1).setValue(draft.getId());
  }
  Logger.log('Drafts created. Now run startSending() to begin auto-sending.');
}


/** STEP B, run this to start auto-sending in batches, hands-free. */
function startSending() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendBatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendBatch')
    .timeBased()
    .everyMinutes(BATCH_INTERVAL_MIN)
    .create();
  sendBatch(); // send the first batch immediately
  Logger.log('Auto-send started: ' + BATCH_SIZE + ' every ' + BATCH_INTERVAL_MIN + ' min.');
}


/** Sends one batch. The trigger calls this automatically, you don't run it by hand. */
function sendBatch() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const head  = data[0];
  const col   = name => head.indexOf(name);

  let sent = 0;
  for (let i = 1; i < data.length && sent < BATCH_SIZE; i++) {
    const row = data[i];
    if (row[col('Status')] !== 'DRAFT') continue;
    const draftId = row[col('DraftId')];
    if (!draftId) continue;

    try {
      const draft = GmailApp.getDraft(draftId);
      draft.send();
      sheet.getRange(i + 1, col('Status') + 1).setValue('SENT');
      sheet.getRange(i + 1, col('SentAt') + 1).setValue(new Date());
      sent++;
    } catch (e) {
      Logger.log('Could not send row ' + (i + 1) + ': ' + e);
    }
  }

  const remaining = data.slice(1).some(r => r[col('Status')] === 'DRAFT');
  if (!remaining) {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'sendBatch') ScriptApp.deleteTrigger(t);
    });
    Logger.log('All emails sent.');
  }
}
