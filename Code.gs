/**
 * DOWN THE RABBIT HOLE — sponsor response tracker
 * Receives partner submissions from the sponsor site and logs them to Google Sheets.
 *
 * SETUP (once):
 *   1. Paste this whole file into Apps Script (replace everything in Code.gs).
 *   2. Edit the three CONFIG values below.
 *   3. Run > setupSheet  (authorize when asked). This builds the headers and formatting.
 *   4. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone > Deploy.
 *   5. Copy the /exec URL into ENDPOINT inside index.html (one line, near the bottom).
 *
 * After ANY edit to this file you must run Deploy > Manage deployments > edit > Version: New version > Deploy,
 * or the live site keeps talking to the old code.
 */

const CONFIG = {
  SHEET_ID:     '11ds8oKoVf4FWrpqZJ4OBiPK26pMqxzPabKKFpW8yf8Y',
  SHEET_NAME:   'Sponsors',
  NOTIFY_EMAIL: 'amechi@19keys.com'   // change to whoever should get the instant alert; comma-separate for several
};

const HEADERS = [
  'Timestamp','Status','Owner','Brand','Contact Name','Email','Phone',
  'Partnership Type','Level','Est. Value','Portals','Areas','In-Kind Offered',
  'What They Want To Create','Follow-Up Date','Source Page','Referrer','Device','Submission ID'
];

const LEVEL_VALUE = {
  'Presenting Partner': 25000,
  'Signature Partner':  10000,
  'Experience Partner':  5000,
  'Supporting Partner':  2500,
  'In-kind Partner':        0,
  'Hybrid Partner':         0
};

const STATUS_OPTIONS = ['New','Contacted','In Conversation','Verbal Yes','Contract Sent','Signed','Paid','Fulfilled','Passed'];

/** Receives the form submission from the website. */
function doPost(e) {
  try {
    const data = parseIncoming_(e);
    if (!data.brand && !data.email) return json_({ ok: false, error: 'Empty submission' });
    if (data['bot-field']) return json_({ ok: true, skipped: 'bot' });

    const sheet = getSheet_();
    const id = 'RH-' + Utilities.formatDate(new Date(), 'America/Chicago', 'yyyyMMdd-HHmmss');
    const level = data.level || '';

    sheet.appendRow([
      new Date(),
      'New',
      '',
      data.brand || '',
      data.name || '',
      data.email || '',
      data.phone || '',
      data.type || '',
      level,
      LEVEL_VALUE[level] || '',
      data.portals || '',
      data.categories || '',
      data.inkind || '',
      data.notes || '',
      '',
      data.page || '',
      data.referrer || '',
      data.device || '',
      id
    ]);

    notify_(data, id);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Health check. Open the /exec URL in a browser to confirm the deployment is live. */
function doGet() {
  return json_({ ok: true, service: 'Down the Rabbit Hole sponsor tracker', time: new Date().toISOString() });
}

/** Run this once from the editor to build the sheet. Safe to re-run. */
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
       .setBackground('#211810').setFontColor('#F7F1E7').setFontWeight('bold')
       .setFontSize(10).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);

  const widths = [150,120,110,170,150,210,130,130,170,110,260,260,260,320,120,200,160,200,160];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('J2:J').setNumberFormat('$#,##0');
  sheet.getRange('O2:O').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('A2:S').setVerticalAlignment('top').setWrap(false);
  sheet.getRange('K2:N').setWrap(true);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build();
  sheet.getRange('B2:B').setDataValidation(statusRule);

  const rules = [];
  const colorMap = { 'Signed': '#C6EFCE', 'Paid': '#A9D08E', 'Passed': '#F2F2F2', 'New': '#FCE4D6', 'Verbal Yes': '#FFF2CC' };
  Object.keys(colorMap).forEach(k => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="' + k + '"')
      .setBackground(colorMap[k])
      .setRanges([sheet.getRange('A2:S1000')]).build());
  });
  sheet.setConditionalFormatRules(rules);

  // Summary tab
  let sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  sum.clear();
  sum.getRange('A1').setValue('DOWN THE RABBIT HOLE — PIPELINE').setFontWeight('bold').setFontSize(14);
  const S = "'" + CONFIG.SHEET_NAME + "'";
  sum.getRange('A3:B10').setValues([
    ['Submissions',            '=COUNTA(' + S + '!D2:D)'],
    ['Monetary',               '=COUNTIF(' + S + '!H2:H,"Monetary")'],
    ['In-kind',                '=COUNTIF(' + S + '!H2:H,"In-kind")'],
    ['Hybrid',                 '=COUNTIF(' + S + '!H2:H,"Hybrid")'],
    ['Pipeline value',         '=SUM(' + S + '!J2:J)'],
    ['Signed + Paid value',    '=SUMIFS(' + S + '!J2:J,' + S + '!B2:B,"Signed")+SUMIFS(' + S + '!J2:J,' + S + '!B2:B,"Paid")'],
    ['Needs first contact',    '=COUNTIF(' + S + '!B2:B,"New")'],
    ['Days to Oct 2 close',    '=DATEDIF(TODAY(),DATE(2026,10,2),"D")']
  ]);
  sum.getRange('A3:A10').setFontWeight('bold');
  sum.getRange('B7:B8').setNumberFormat('$#,##0');
  sum.setColumnWidth(1, 200); sum.setColumnWidth(2, 160);

  SpreadsheetApp.flush();
  return 'Sheet ready: ' + ss.getUrl();
}

/** Accepts JSON body, form-encoded body, or URL parameters. */
function parseIncoming_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (ignore) {}
  }
  const out = {};
  if (e && e.parameter) Object.keys(e.parameter).forEach(k => out[k] = e.parameter[k]);
  return out;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { setupSheet(); sheet = ss.getSheetByName(CONFIG.SHEET_NAME); }
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function notify_(d, id) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  const subject = 'Rabbit Hole partner: ' + (d.brand || 'Unnamed brand') + (d.level ? ' — ' + d.level : '');
  const rows = [
    ['Brand', d.brand], ['Contact', d.name], ['Email', d.email], ['Phone', d.phone],
    ['Type', d.type], ['Level', d.level], ['Portals', d.portals], ['Areas', d.categories],
    ['In-kind', d.inkind], ['Wants to create', d.notes], ['Submission', id]
  ].filter(r => r[1]).map(r =>
    '<tr><td style="padding:6px 14px 6px 0;color:#8a7a6a;font:600 11px/1.4 Arial;letter-spacing:.08em;text-transform:uppercase;vertical-align:top">' +
    r[0] + '</td><td style="padding:6px 0;font:14px/1.5 Arial;color:#211810">' + String(r[1]).replace(/</g, '&lt;') + '</td></tr>'
  ).join('');

  const html =
    '<div style="background:#120D08;padding:22px 24px;color:#F7F1E7;font:600 13px Arial;letter-spacing:.18em">NEW PARTNER SUBMISSION</div>' +
    '<div style="padding:22px 24px;background:#F7F1E7"><table>' + rows + '</table>' +
    '<p style="font:14px Arial"><a href="https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '/edit" ' +
    'style="background:#B4471C;color:#F7F1E7;padding:12px 20px;border-radius:8px;text-decoration:none;display:inline-block;font-weight:bold">Open the tracker</a></p></div>';

  MailApp.sendEmail({ to: CONFIG.NOTIFY_EMAIL, subject: subject, htmlBody: html, replyTo: d.email || undefined });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
