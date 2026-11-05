/**
 * Intergalactic Ad Pitch: Round 1 quiz backend (Google Apps Script)
 * ---------------------------------------------------------------
 * - Keeps the ANSWER KEY private (it is never sent to the quiz page).
 * - Scores every team on the server, so inspect element can't fake a score.
 * - Writes results to this Google Sheet. The "Leaderboard" tab is only
 *   visible to you (and anyone you share the Sheet with).
 *
 * Setup: see README.md. Run setup() once, then Deploy → Web app.
 */

// Correct option for each question: 0 = A, 1 = B, 2 = C, 3 = D.
// Must stay in the same order as QUESTIONS in index.html.
const ANSWER_KEY = [1, 3, 0, 1, 2, 0, 2, 3, 0, 2, 1, 3, 0, 1, 3, 0, 2, 1, 0, 2];
const POINTS_PER_QUESTION = 100;

const RESPONSES = 'Responses';
const LEADERBOARD = 'Leaderboard';
// Columns in Responses: A Started at | B Team | C Token | D Status | E Score | F Correct | G Finished at | H Answers

/** Run this ONCE from the Apps Script editor (select "setup" → Run). */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = ss.getSheetByName(RESPONSES) || ss.insertSheet(RESPONSES);
  r.clear();
  r.appendRow(['Started at', 'Team', 'Token', 'Status', 'Score', 'Correct', 'Finished at', 'Answers']);
  r.setFrozenRows(1);
  r.getRange('1:1').setFontWeight('bold');

  const l = ss.getSheetByName(LEADERBOARD) || ss.insertSheet(LEADERBOARD, 0);
  l.clear();
  l.getRange('A1').setFormula(
    '=QUERY(' + RESPONSES + '!A:G,"select B, E, F, G where D = \'Finished\' order by E desc, G asc ' +
    'label B \'Team\', E \'Score\', F \'Correct (of 20)\', G \'Finished at\'",1)'
  );
  l.setFrozenRows(1);
  l.getRange('1:1').setFontWeight('bold');
  l.setColumnWidth(1, 240);
}

function doGet() {
  return json_({ ok: true, message: 'Quiz backend is running.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.action === 'register') return json_(register_(data));
    if (data.action === 'submit') return json_(submit_(data));
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: 'Server busy, please try again.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function register_(data) {
  let team = String(data.team || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (team.length < 2) return { ok: false, error: 'Enter your team name.' };
  if (/^[=+\-@]/.test(team)) team = "'" + team; // stop spreadsheet formula injection

  const sheet = sheet_();
  const last = sheet.getLastRow();
  if (last > 1) {
    const names = sheet.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0]).toLowerCase().replace(/^'/, ''));
    if (names.indexOf(team.toLowerCase().replace(/^'/, '')) !== -1) return { ok: false, error: 'taken' };
  }
  const token = Utilities.getUuid();
  sheet.appendRow([new Date(), team, token, 'Started', '', '', '', '']);
  return { ok: true, token: token };
}

function submit_(data) {
  const token = String(data.token || '');
  const answers = data.answers;
  if (!token) return { ok: false, error: 'Missing team.' };

  const sheet = sheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { ok: false, error: 'Team not found.' };
  const tokens = sheet.getRange(2, 3, last - 1, 1).getValues();
  let row = -1;
  for (let i = 0; i < tokens.length; i++) if (tokens[i][0] === token) { row = i + 2; break; }
  if (row === -1) return { ok: false, error: 'Team not found. Ask an organiser.' };

  const rec = sheet.getRange(row, 4, 1, 3).getValues()[0]; // Status, Score, Correct
  if (rec[0] === 'Finished') return { ok: true, score: rec[1], correct: rec[2] }; // no second attempts

  if (!Array.isArray(answers) || answers.length !== ANSWER_KEY.length) return { ok: false, error: 'Incomplete answers.' };
  let correct = 0;
  for (let i = 0; i < ANSWER_KEY.length; i++) if (Number(answers[i]) === ANSWER_KEY[i]) correct++;
  const score = correct * POINTS_PER_QUESTION;
  const letters = answers.map(a => 'ABCD'.charAt(Number(a)) || '?').join(' ');

  sheet.getRange(row, 4, 1, 5).setValues([['Finished', score, correct, new Date(), letters]]);
  return { ok: true, score: score, correct: correct };
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(RESPONSES);
  if (!s) { setup(); return ss.getSheetByName(RESPONSES); }
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
