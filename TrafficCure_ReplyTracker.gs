/*** ============ REPLY TRACKER, add-on for the TrafficCure mailer ============ ***
 *  Paste this as a SECOND file in the SAME Apps Script project as the mailer.
 *  It reads Sir's own Gmail to see who replied. No pixels, no tracking links.
 *
 *  What it does:
 *   - For every row marked SENT, it finds the email thread in Gmail
 *   - Checks if that person sent a reply back
 *   - Fills in three columns: Replied, ReplyAt, ReplySnippet
 *     (it creates those columns automatically if they are missing)
 ***/

const REPLY_SHEET_NAME       = 'Sheet1';
const REPLY_CHECK_EVERY_HOURS = 6;   // how often the auto-check runs


/** RUN THIS to check replies once, right now. Safe to run as often as you like. */
function checkReplies() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(REPLY_SHEET_NAME);
  ensureReplyColumns_(sheet);

  const data = sheet.getDataRange().getValues();
  const head = data[0];
  const col  = name => head.indexOf(name);

  let checked = 0, replied = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col('Status')] !== 'SENT') continue;        // only look at ones we sent
    if (row[col('Replied')] === 'YES') { replied++; continue; } // already known, skip

    const email   = String(row[col('Email')] || '').trim().toLowerCase();
    const subject = String(row[col('Subject')] || '').trim();
    if (!email) continue;
    checked++;

    // Find the thread: use the stored ThreadId if we have one, else search Gmail.
    let thread = null;
    const storedId = row[col('ThreadId')];
    if (storedId) {
      try { thread = GmailApp.getThreadById(storedId); } catch (e) { thread = null; }
    }
    if (!thread) {
      const query   = 'to:' + email + ' subject:' + JSON.stringify(subject);
      const threads = GmailApp.search(query, 0, 3);
      if (threads.length) {
        thread = threads[0];
        sheet.getRange(i + 1, col('ThreadId') + 1).setValue(thread.getId());
      }
    }

    if (!thread) {
      sheet.getRange(i + 1, col('Replied') + 1).setValue('NOT FOUND');
      continue;
    }

    // Look through the thread for a message that came FROM the recipient.
    const messages = thread.getMessages();
    let replyDate = null, snippet = '';
    for (let m = 0; m < messages.length; m++) {
      const from = String(messages[m].getFrom() || '').toLowerCase();
      if (from.indexOf(email) !== -1) {                 // this message is from them = a reply
        replyDate = messages[m].getDate();
        snippet   = String(messages[m].getPlainBody() || '')
                      .replace(/\s+/g, ' ').trim().slice(0, 200);
      }
    }

    if (replyDate) {
      sheet.getRange(i + 1, col('Replied') + 1).setValue('YES');
      sheet.getRange(i + 1, col('ReplyAt') + 1).setValue(replyDate);
      sheet.getRange(i + 1, col('ReplySnippet') + 1).setValue(snippet);
      replied++;
    } else {
      sheet.getRange(i + 1, col('Replied') + 1).setValue('NO');
    }
  }

  Logger.log('Checked ' + checked + ' sent emails. Replies so far: ' + replied + '.');
}


/** RUN THIS once to keep checking automatically every few hours, hands-free. */
function startReplyTracking() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkReplies') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkReplies')
    .timeBased()
    .everyHours(REPLY_CHECK_EVERY_HOURS)
    .create();
  checkReplies(); // do one check right away
  Logger.log('Auto reply-tracking started: every ' + REPLY_CHECK_EVERY_HOURS + ' hours.');
}


/** RUN THIS to stop the automatic checking. */
function stopReplyTracking() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkReplies') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Auto reply-tracking stopped.');
}


/** Helper: makes sure the Replied, ReplyAt, ReplySnippet, ThreadId columns exist. */
function ensureReplyColumns_(sheet) {
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['ThreadId', 'Replied', 'ReplyAt', 'ReplySnippet'].forEach(name => {
    if (head.indexOf(name) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name);
    }
  });
}
