/***** CONFIG *****/
const SHEET_ID = '{{INSERT YOUR SHEET ID HERE}}'; // from your URL
const ACTIONS_SHEET_NAME = 'Email Actions';
const APP_TITLE = 'Gmail Virtual Assistant';
const ARCHIVE_QUERY = 'in:inbox newer_than:2d';         // Only last 7 days
const SENT_QUERY    = 'in:sent newer_than:2d';           // Only last 7 days
const MAX_THREADS_PER_RUN = 500;                          // safety cap

/***** MAIN: Auto-archive last 7 days of inbox *****/
function autoArchiveLast7Days() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ensureLogSheet_(ss);
  const rulesSheet = ss.getSheetByName('Rules');
  if (!rulesSheet) throw new Error('Missing "Rules" sheet tab.');

  // Load rule strings from column B (Rule Condition Contains)
  const rules = rulesSheet
    .getRange(2, 2, rulesSheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(v => v && v.toString().trim() !== '');

  Logger.log(`✅ Loaded ${rules.length} rules from "Rules" sheet.`);
  const now = new Date();
  const lock = LockService.getScriptLock();
  lock.tryLock(20 * 1000);

  let totalThreads = 0;
  let matchedThreads = 0;
  let archivedThreads = 0;
  const rows = [];

  try {
    const threads = GmailApp.search(ARCHIVE_QUERY, 0, MAX_THREADS_PER_RUN);
    totalThreads = threads.length;
    Logger.log(`📬 Found ${totalThreads} inbox threads from last 7 days.`);

    const chunkSize = 50;
    for (let i = 0; i < threads.length; i += chunkSize) {
      const chunk = threads.slice(i, i + chunkSize);
      Logger.log(`🔹 Processing threads ${i + 1}–${Math.min(i + chunk.length, totalThreads)}...`);
      const toArchive = [];

      chunk.forEach(thread => {
        const msgs = thread.getMessages();
        let shouldArchive = false;
        let matchedRule = '';

        for (const msg of msgs) {
          const from = (msg.getFrom() || '').toLowerCase();
          const subj = (msg.getSubject() || '').toLowerCase();
          const body = (msg.getPlainBody() || '').toLowerCase();

          for (const rule of rules) {
            const ruleStr = rule.toString().toLowerCase();
            if (from.includes(ruleStr) || subj.includes(ruleStr) || body.includes(ruleStr)) {
              shouldArchive = true;
              matchedRule = rule;
              Logger.log(`✅ Match found: "${rule}" in email subject "${subj.slice(0, 60)}..."`);
              break;
            }
          }
          if (shouldArchive) break;
        }

        if (shouldArchive) {
          matchedThreads++;
          toArchive.push(thread);

          // Mark thread as read before archiving
          thread.markRead();

          const firstMsg = msgs[0];
          const subj = safeSubject_(firstMsg.getSubject());
          rows.push([
            Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
            'Marked as Read + Archived',
            subj,
            `Matched rule: "${matchedRule}"`
          ]);
        }
      });

      if (toArchive.length > 0) {
        GmailApp.moveThreadsToArchive(toArchive);
        archivedThreads += toArchive.length;
        Logger.log(`📦 Marked as read & archived ${toArchive.length} threads in this batch.`);
      } else {
        Logger.log(`⏩ No matches found in this batch.`);
      }
    }

    if (rows.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
      Logger.log(`🧾 Logged ${rows.length} archived entries to "Email Actions" sheet.`);
    } else {
      Logger.log('⚠️ No threads met any rule conditions this run.');
    }
  } catch (err) {
    Logger.log(`❌ Error: ${err.message}`);
    throw err;
  } finally {
    lock.releaseLock();
  }

  Logger.log(`✅ Run complete.`);
  Logger.log(`Summary → Total Threads: ${totalThreads}, Matched: ${matchedThreads}, Archived: ${archivedThreads}`);
}




/***** Build Sent Mail Summary (last 7d) for the Web App *****/
function getSentSummary() {
  try {
    Logger.log('📨 Starting getSentSummary...');
    const threads = GmailApp.search(SENT_QUERY, 0, 200); // keep light for speed
    const items = [];

    threads.forEach(th => {
      const msgs = th.getMessages();
      if (!msgs || !msgs.length) return;

      msgs.forEach(m => {
        // defensive: ensure GmailMessage API exists
        if (!m || typeof m.getFrom !== 'function') return;

        const from = (m.getFrom() || '').toLowerCase();
        const myEmail = Session.getActiveUser().getEmail().toLowerCase();

        // Only include messages you actually sent
        if (!from.includes(myEmail)) return;

        const date = m.getDate();
        const subject = safeSubject_(m.getSubject());
        const to = (m.getTo() || '').split(',').map(s => s.trim()).filter(Boolean);
        const cc = (m.getCc() || '').split(',').map(s => s.trim()).filter(Boolean);
        const bcc = (m.getBcc() || '').split(',').map(s => s.trim()).filter(Boolean);
        const recipients = [...to, ...cc, ...bcc];
        const domains = recipients
          .map(r => r.split('@')[1] || '')
          .filter(Boolean)
          .map(d => d.toLowerCase());

        items.push({
          date: date,
          iso: date.toISOString(),
          subject: subject,
          recipients: recipients,
          domains: [...new Set(domains)],
          threadId: th.getId(),
          snippet: (m.getPlainBody() || '').slice(0, 200)
        });
      });
    });

    Logger.log(`✅ getSentSummary collected ${items.length} messages`);

    return {
      generatedAt: new Date().toISOString(),
      totals: { sentMessages: items.length },
      categories: { byDomain: {}, byDay: {}, byTopic: {} },
      items
    };
  } catch (err) {
    Logger.log(`❌ getSentSummary error: ${err.message}`);
    return {
      error: true,
      message: err.message,
      totals: { sentMessages: 0 },
      categories: {},
      items: []
    };
  }
}



/***** Web App *****/
function doGet() {
  const t = HtmlService.createTemplateFromFile('WebApp');
  t.title = APP_TITLE;
  return t.evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // so you can embed if needed
}

/***** Utilities *****/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureLogSheet_(ss) {
  let sh = ss.getSheetByName(ACTIONS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ACTIONS_SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([['Date Time','Action Taken','Email Subject','Rule / Reason Action Taken']]);
    sh.setFrozenRows(1);
  } else {
    // ensure headers exist
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 4).setValues([['Date Time','Action Taken','Email Subject','Rule / Reason Action Taken']]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function safeSubject_(s) {
  if (!s) return '(no subject)';
  // guard against formula injection into Sheets
  const startsWithFormulaChar = /^[=+\-@']/;
  return startsWithFormulaChar.test(s) ? "'" + s : s;
}

// Convert objects with arrays of complex items to client-safe
function sanitizeForClient_(obj) {
  const out = {};
  Object.keys(obj).forEach(k => {
    const entry = obj[k];
    out[k] = {
      count: entry.count,
      items: entry.items.map(it => ({
        iso: it.iso,
        subject: it.subject,
        recipients: it.recipients,
        domains: it.domains,
        snippet: it.snippet,
        threadId: it.threadId,
      }))
    };
  });
  return out;
}

/***** Sheet Log Reader (optional in UI) *****/
function getActionLog(limit) {
  try {
    Logger.log('📋 Loading action log...');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ensureLogSheet_(ss);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const start = Math.max(2, last - (limit || 200) + 1);
    const rng = sh.getRange(start, 1, last - start + 1, 4).getValues();
    const data = rng.map(r => ({ dateTime: r[0], action: r[1], subject: r[2], reason: r[3] })).reverse();
    Logger.log(`✅ getActionLog returned ${data.length} rows`);
    return data;
  } catch (err) {
    Logger.log(`❌ getActionLog error: ${err.message}`);
    return [];
  }
}


/***** Menu and Trigger Helpers *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gmail Assistant')
    .addItem('Run Auto-Archive (last 7 days)', 'autoArchiveLast7Days')
    .addItem('Open Web App (preview)', 'openWebAppInSidebar_')
    .addSeparator()
    .addItem('Install Daily Trigger', 'installDailyTrigger')
    .addToUi();
}

function openWebAppInSidebar_() {
  const html = HtmlService.createTemplateFromFile('WebApp').evaluate().setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function installDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'autoArchiveLast7Days') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoArchiveLast7Days')
    .timeBased()
    .everyDays(1)
    .atHour(7) // 7am your script timezone
    .create();
}
