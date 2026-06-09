/*** ===================== TrafficCure FOLLOW-UP + REPLY ALERT =====================
 *  Paste this as a file in the SAME Apps Script project as the mailer.
 *
 *  What one cycle does, automatically, for every contact already marked SENT:
 *    1. Looks in Sir's Gmail to see if that person has REPLIED.
 *       - If they replied  -> sends Sir an alert email so he can handle it himself,
 *                             and that person is NEVER sent a follow-up.
 *    2. If they have NOT replied and enough days have passed since the first email,
 *       it sends their personalised follow-up (the FollowUpSubject / FollowUpBody
 *       columns in the sheet) and marks the row so it is never followed up twice.
 *
 *  No tracking pixels, no tracking links. It only reads Sir's own mailbox.
 *
 *  HOW TO RUN (non-technical):
 *    - Run  startFollowUpAutomation  ONCE. It checks now, then re-checks on its own.
 *    - To check by hand any time, run  runFollowUpCycle.
 *    - To turn the automation off, run  stopFollowUpAutomation.
 ******************************************************************************/

/*** ============================ SETTINGS ============================ ***/
const FU_SHEET_NAME        = 'Sheet1';
const FOLLOWUP_AFTER_DAYS  = 4;     // days of silence before a follow-up is sent
const CHECK_EVERY_HOURS    = 6;     // how often the automation re-checks
const FOLLOWUP_MAX_PER_RUN = 10;    // safety cap: max follow-ups sent in one cycle
// Where reply alerts go. Leave '' to send them to Sir's own Gmail (the signed-in
// account). Or put any address, e.g. 'team@leptonmaps.com'.
const ALERT_EMAIL          = '';
/*** ================================================================== ***/


/** MAIN. Run this to do one full pass: detect replies, alert, send due follow-ups. */
function runFollowUpCycle() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(FU_SHEET_NAME);
  if (!sheet) { Logger.log('Sheet "' + FU_SHEET_NAME + '" not found.'); return; }
  ensureFollowUpColumns_(sheet);

  const data = sheet.getDataRange().getValues();
  const head = data[0];
  const col  = name => head.indexOf(name);

  const alertTo = ALERT_EMAIL || Session.getActiveUser().getEmail();
  const now     = new Date();

  let newReplies = 0, followUpsSent = 0, waiting = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col('Status')] !== 'SENT') continue;          // only people we have emailed

    const email = String(row[col('Email')] || '').trim();
    if (!email) continue;
    const emailLc = email.toLowerCase();
    const name    = String(row[col('Name')] || '').trim();
    const company = String(row[col('Company')] || '').trim();
    const subject = String(row[col('Subject')] || '').trim();

    // ---- 1. Has this person replied? ----
    let replied = row[col('Replied')] === 'YES';
    if (!replied) {
      const found = findReply_(emailLc, subject, row[col('ThreadId')]);
      if (found.threadId && !row[col('ThreadId')]) {
        sheet.getRange(i + 1, col('ThreadId') + 1).setValue(found.threadId);
      }
      if (found.replyDate) {
        replied = true;
        sheet.getRange(i + 1, col('Replied') + 1).setValue('YES');
        sheet.getRange(i + 1, col('ReplyAt') + 1).setValue(found.replyDate);
        sheet.getRange(i + 1, col('ReplySnippet') + 1).setValue(found.snippet);
      }
    }

    // ---- 2a. Replied -> alert Sir once, never follow up. ----
    if (replied) {
      if (row[col('Notified')] !== 'YES') {
        sendReplyAlert_(alertTo, name, company, email,
                        String(row[col('ReplySnippet')] || ''));
        sheet.getRange(i + 1, col('Notified') + 1).setValue('YES');
        newReplies++;
      }
      continue; // a reply means we stop here, no follow-up ever
    }

    // ---- 2b. No reply -> send the follow-up if it is due and not already sent. ----
    if (row[col('FollowUpStatus')] === 'SENT') continue;       // already followed up
    if (followUpsSent >= FOLLOWUP_MAX_PER_RUN) { waiting++; continue; }

    const sentAt = row[col('SentAt')];
    if (!(sentAt instanceof Date)) {
      // No send timestamp on record, so we cannot tell how long it has been waiting.
      sheet.getRange(i + 1, col('FollowUpStatus') + 1).setValue('NO SentAt');
      continue;
    }
    const daysSince = (now - sentAt) / (1000 * 60 * 60 * 24);
    if (daysSince < FOLLOWUP_AFTER_DAYS) { waiting++; continue; } // too soon

    const fuSubject = String(row[col('FollowUpSubject')] || '').trim();
    const fuBody    = String(row[col('FollowUpBody')] || '').trim();
    if (!fuSubject || !fuBody) {
      sheet.getRange(i + 1, col('FollowUpStatus') + 1).setValue('NO CONTENT');
      continue;
    }

    try {
      GmailApp.sendEmail(email, fuSubject, fuBody,
                         { htmlBody: fuBody.replace(/\n/g, '<br>') });
      sheet.getRange(i + 1, col('FollowUpStatus') + 1).setValue('SENT');
      sheet.getRange(i + 1, col('FollowUpAt') + 1).setValue(new Date());
      followUpsSent++;
    } catch (e) {
      Logger.log('Could not send follow-up to ' + email + ': ' + e);
      sheet.getRange(i + 1, col('FollowUpStatus') + 1).setValue('ERROR');
    }
  }

  Logger.log('Cycle done. New replies alerted: ' + newReplies +
             '. Follow-ups sent: ' + followUpsSent +
             '. Still waiting (no reply, not yet due / capped): ' + waiting + '.');
}


/** Run ONCE to keep checking and following up automatically, hands-free. */
function startFollowUpAutomation() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runFollowUpCycle') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runFollowUpCycle')
    .timeBased()
    .everyHours(CHECK_EVERY_HOURS)
    .create();
  runFollowUpCycle(); // do one pass right now
  Logger.log('Follow-up automation started: every ' + CHECK_EVERY_HOURS + ' hours.');
}


/** Run to turn the automation off. */
function stopFollowUpAutomation() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runFollowUpCycle') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Follow-up automation stopped.');
}


/* ----------------------------- helpers ----------------------------- */

/** Finds whether the contact replied. Returns {threadId, replyDate, snippet}. */
function findReply_(emailLc, subject, storedThreadId) {
  let thread = null;
  if (storedThreadId) {
    try { thread = GmailApp.getThreadById(storedThreadId); } catch (e) { thread = null; }
  }
  if (!thread) {
    const query   = 'to:' + emailLc + ' subject:' + JSON.stringify(subject);
    const threads = GmailApp.search(query, 0, 3);
    if (threads.length) thread = threads[0];
  }
  if (!thread) return { threadId: '', replyDate: null, snippet: '' };

  let replyDate = null, snippet = '';
  const messages = thread.getMessages();
  for (let m = 0; m < messages.length; m++) {
    const from = String(messages[m].getFrom() || '').toLowerCase();
    if (from.indexOf(emailLc) !== -1) {            // a message FROM them is a reply
      replyDate = messages[m].getDate();
      snippet   = String(messages[m].getPlainBody() || '')
                    .replace(/\s+/g, ' ').trim().slice(0, 250);
    }
  }
  return { threadId: thread.getId(), replyDate: replyDate, snippet: snippet };
}


/** Emails Sir an alert that someone replied. */
function sendReplyAlert_(alertTo, name, company, email, snippet) {
  const subject = 'Reply received: ' + (name || email) +
                  (company ? ' (' + company + ')' : '');
  const body =
    name + ' from ' + company + ' has replied to the TrafficCure outreach.\n\n' +
    'Email: ' + email + '\n\n' +
    'What they wrote (start of their reply):\n' +
    (snippet || '(no preview available)') + '\n\n' +
    'No follow-up will be sent to this person. Please reply to them directly from ' +
    'the inbox.';
  try {
    GmailApp.sendEmail(alertTo, subject, body);
  } catch (e) {
    Logger.log('Could not send reply alert for ' + email + ': ' + e);
  }
}


/** Makes sure the tracking / follow-up status columns exist; creates any missing. */
function ensureFollowUpColumns_(sheet) {
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ['ThreadId', 'Replied', 'ReplyAt', 'ReplySnippet',
   'FollowUpStatus', 'FollowUpAt', 'Notified'].forEach(name => {
    if (head.indexOf(name) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name);
    }
  });
}
