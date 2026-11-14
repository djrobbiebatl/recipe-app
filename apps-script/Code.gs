/**
 * Recipe API — paste into Extensions > Apps Script of the recipes Google Sheet.
 * Then: Project Settings > Script properties > add PASSCODE = your shared passcode,
 * and Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The passcode is the only gate, so make it a real one (12+ characters).
 */
const SHEET_NAME = 'Recipes';
const COLS = ['id', 'title', 'category', 'sourceName', 'sourceUrl', 'servings', 'summary',
              'image', 'ingredients', 'steps', 'addedBy', 'added', 'favorite', 'prepTime', 'totalTime', 'nutrition', 'notes'];

function doGet() { return out_({ ok: true, msg: 'Recipe API is running' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass || String(req.key) !== pass) return out_({ ok: false, error: 'auth' });

  if (req.action === 'list') return out_({ ok: true, recipes: list_() });
  if (req.action === 'save') return save_(req.recipe, req.name);
  if (req.action === 'import') return import_(req.url);
  if (req.action === 'delete') return delete_(req.id);
  if (req.action === 'scan') return scan_(req.data, req.mime);
  if (req.action === 'upload') return upload_(req.data, req.mime);
  if (req.action === 'parse') return parseText_(req.text);
  if (req.action === 'intakeadd') return intakeAddAction_(req.item, req.address);
  if (req.action === 'intakesync') return intakeListAction_();
  if (req.action === 'intakecount') return intakeCountAction_();
  if (req.action === 'intakeget') return intakeGetAction_(req.id);
  if (req.action === 'intakemark') return intakeMarkAction_(req.id, req.status);
  if (req.action === 'drivelist') return driveList_();
  if (req.action === 'driveimport') return driveImport_(req.id);
  return out_({ ok: false, error: 'unknown action' });
}

function delete_(id) {
  if (!id) return out_({ ok: false, error: 'missing id' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    if (n < 1) return out_({ ok: true });
    const ids = sh.getRange(2, 1, n, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

// ---------- email intake: recipes forwarded to a dedicated intake Gmail account wait in the "Intake" tab ----------
// This script has NO access to any mailbox. The separate intake account runs apps-script/Intake.gs, which pushes emails here.
const INTAKE_COLS = ['id', 'from', 'subject', 'date', 'text', 'status', 'added'];

function intakeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Intake');
  if (!sh) {
    sh = ss.insertSheet('Intake');
    sh.getRange(1, 1, 1, INTAKE_COLS.length).setNumberFormat('@').setValues([INTAKE_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function intakeRows_() {
  const sh = intakeSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, INTAKE_COLS.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0]), from: String(r[1]), subject: String(r[2]), date: String(r[3]), text: String(r[4]), status: String(r[5] || 'pending') };
  });
}

/**
 * Called by the small "Intake.gs" script that lives in the SEPARATE intake Google account. It pushes each new email
 * here (with the shared passcode). This script never touches any mailbox itself.
 */
function intakeAddAction_(item, address) {
  if (!item || !item.id) return out_({ ok: false, error: 'bad email' });
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(address || ''))) PropertiesService.getScriptProperties().setProperty('INTAKE_ADDRESS', String(address).slice(0, 120));
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const rows = intakeRows_();
    const id = String(item.id).slice(0, 80);
    if (rows.some(function (r) { return r.id === id; })) return out_({ ok: true, duplicate: true });
    if (rows.filter(function (r) { return r.status === 'pending'; }).length >= 150) return out_({ ok: false, error: 'queue is full' });
    const text = composeEmail_(String(item.subject || '').slice(0, 300), String(item.plain || '').slice(0, 60000)).slice(0, 40000);
    const empty = text.replace(/\s/g, '').length < 20;
    const row = [id, String(item.from || '').slice(0, 200), text.split('\n')[0].slice(0, 200), String(item.date || new Date().toISOString()).slice(0, 40),
                 text, empty ? 'skipped' : 'pending', new Date().toISOString()];
    const sh = intakeSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, 1, INTAKE_COLS.length).setNumberFormat('@').setValues([row]);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

function intakeListAction_() {
  const items = intakeRows_().filter(function (r) { return r.status === 'pending'; }).reverse().map(function (r) {
    return { id: r.id, from: r.from, subject: r.subject, date: r.date, snippet: r.text.split('\n').slice(1).join(' ').replace(/\s+/g, ' ').slice(0, 180) };
  });
  return out_({ ok: true, address: PropertiesService.getScriptProperties().getProperty('INTAKE_ADDRESS') || '', items: items });
}

function intakeCountAction_() {
  return out_({ ok: true, count: intakeRows_().filter(function (r) { return r.status === 'pending'; }).length });
}

function intakeGetAction_(id) {
  const r = intakeRows_().filter(function (x) { return x.id === String(id); })[0];
  if (!r) return out_({ ok: false, error: 'That email isn\'t in the queue any more' });
  return out_({ ok: true, text: r.text.slice(0, 20000), fileUrl: '', from: r.from, recipe: parseScan_(r.text) });
}

function intakeMarkAction_(id, status) {
  if (['posted', 'dismissed', 'pending'].indexOf(status) < 0) return out_({ ok: false, error: 'bad status' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = intakeRows_().filter(function (x) { return x.id === String(id); })[0];
    if (r) intakeSheet_().getRange(r.row, 6).setValue(status);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

// ---------- import from Drive: a "Recipe Inbox" folder to pick from, or a pasted Drive/Docs link ----------
const INBOX_NAME = 'Recipe Inbox';

function inbox_() {
  const it = DriveApp.getFoldersByName(INBOX_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(INBOX_NAME);
}

function isRecipeFile_(f) {
  const m = f.getMimeType();
  return /^(image\/(jpeg|png|webp|gif)|application\/pdf|application\/msword|application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document|text\/plain|message\/rfc822|application\/vnd\.google-apps\.document)$/.test(m) ||
    /\.eml$/i.test(f.getName());
}

function driveList_() {
  const folder = inbox_();
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext() && files.length < 100) { const f = it.next(); if (isRecipeFile_(f)) files.push(f); }
  files.sort(function (a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
  return out_({ ok: true, folderUrl: folder.getUrl(), folderName: INBOX_NAME, files: files.slice(0, 40).map(function (f) {
    return { id: f.getId(), name: f.getName(), modified: f.getLastUpdated().toISOString() };
  }) });
}

function driveImport_(id) {
  if (!/^[\w-]{10,}$/.test(String(id || ''))) return out_({ ok: false, error: 'That doesn\'t look like a Drive file' });
  let file;
  try { file = DriveApp.getFileById(id); }
  catch (e) { return out_({ ok: false, error: 'I can\'t open that file. It has to be in your Drive or shared with you.' }); }
  if (file.getMimeType() === 'application/vnd.google-apps.folder') return out_({ ok: false, error: 'That\'s a folder, not a recipe file' });
  try {
    const text = textFromFile_(file);
    if (!text || text.replace(/\s/g, '').length < 20) return out_({ ok: false, error: 'I couldn\'t find any readable text in that file' });
    const r = parseScan_(text);
    r.sourceUrl = file.getUrl(); r.sourceName = 'Import';
    return out_({ ok: true, found: true, text: text.slice(0, 20000), fileUrl: file.getUrl(), recipe: r });
  } catch (e) {
    const msg = String(e && e.message || e);
    if (/Drive is not defined/i.test(msg)) return out_({ ok: false, error: 'The Drive API service isn\'t added to the script yet (Services > + > Drive API)' });
    if (/permission|authori/i.test(msg)) return out_({ ok: false, error: 'The script needs permission to read Drive. Run "authorize" in Apps Script, then redeploy.' });
    return out_({ ok: false, error: 'Couldn\'t read that file (' + msg.slice(0, 120) + ')' });
  }
}

/** File id from a drive.google.com / docs.google.com link, or '' if it isn't one. */
function driveIdFromUrl_(url) {
  const m = String(url).match(/^https?:\/\/(?:drive|docs)\.google\.com\/[^?#]*?\/d\/([\w-]{10,})/i) ||
            String(url).match(/^https?:\/\/(?:drive|docs)\.google\.com\/[^#]*[?&]id=([\w-]{10,})/i);
  return m ? m[1] : '';
}

// ---------- link import (runs here because browsers can't read other sites' pages) ----------
function import_(url) {
  if (!/^https?:\/\/[^\s]+$/i.test(String(url || ''))) return out_({ ok: false, error: 'That doesn\'t look like a web link' });
  const driveId = driveIdFromUrl_(url);
  if (driveId) return driveImport_(driveId);
  let html;
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; RecipeApp/1.0)', 'Accept-Language': 'en-US,en;q=0.9' }
    });
    if (res.getResponseCode() >= 400) return out_({ ok: false, error: 'The site returned error ' + res.getResponseCode() });
    html = res.getContentText();
  } catch (e) {
    const msg = String(e && e.message || e);
    if (/permission|authori/i.test(msg)) return out_({ ok: false, error: 'The script needs permission to read web pages. In Apps Script, run the "authorize" function once, then redeploy.' });
    return out_({ ok: false, error: 'Could not reach that page (' + msg.slice(0, 120) + ')' });
  }
  const r = extract_(html, url);
  return out_({ ok: true, found: r.found, recipe: r.recipe });
}

/**
 * Run this once from the editor (select "authorize" > Run) to grant the permissions for reading web
 * pages and scanning photos/PDFs. Scanning also needs Services (+) > "Drive API" added to this project.
 */
function authorize() {
  UrlFetchApp.fetch('https://example.com');
  const d = DocumentApp.create('authorize-test (safe to delete)');
  DriveApp.getFileById(d.getId()).setTrashed(true);
  Logger.log('Authorized. Now redeploy: Deploy > Manage deployments > edit > New version.');
}

// ---------- kept files: originals of scans and recipe photos live in one Drive folder ----------
function storeFile_(b64, mime, name) {
  const dirs = DriveApp.getFoldersByName('Recipe App Files');
  const folder = dirs.hasNext() ? dirs.next() : DriveApp.createFolder('Recipe App Files');
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  // "Anyone with the link can view" so the app and both of you can open it without Drive sign-in.
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { viewUrl: f.getUrl(), imageUrl: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200' };
}

function upload_(b64, mime) {
  if (!b64 || !/^image\/(jpeg|png|webp|gif)$/.test(String(mime))) return out_({ ok: false, error: 'Please choose a JPEG or PNG photo' });
  try {
    const s = storeFile_(b64, mime, 'recipe-photo-' + Date.now());
    return out_({ ok: true, imageUrl: s.imageUrl, viewUrl: s.viewUrl });
  } catch (e) { return out_({ ok: false, error: 'Upload failed (' + String(e && e.message || e).slice(0, 120) + ')' }); }
}

// ---------- photo / PDF scan: Google Drive turns the file into a Doc (with OCR), we read the text ----------
function scan_(b64, mime) {
  const okMime = /^(image\/(jpeg|png|webp|gif)|application\/pdf|application\/msword|application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document|text\/plain|message\/rfc822)$/;
  if (!b64 || !okMime.test(String(mime))) return out_({ ok: false, error: 'Please use a photo, PDF, Word document (.doc/.docx), text file or email (.eml)' });
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, 'recipe-scan');
    const text = textFromBlob_(blob, mime, '');
    if (!text || text.replace(/\s/g, '').length < 20) return out_({ ok: false, error: 'I couldn\'t find any readable text in that file' });
    let fileUrl = '';
    try { fileUrl = storeFile_(b64, mime, 'recipe-scan-' + Date.now()).viewUrl; } catch (e3) { /* keep going without a link */ }
    return out_({ ok: true, text: text.slice(0, 20000), fileUrl: fileUrl, recipe: parseScan_(text) });
  } catch (e) {
    const msg = String(e && e.message || e);
    if (/Drive is not defined/i.test(msg)) return out_({ ok: false, error: 'The Drive API service isn\'t added to the script yet (Services > + > Drive API)' });
    if (/permission|authori/i.test(msg)) return out_({ ok: false, error: 'The script needs permission to scan. Run "authorize" in Apps Script, then redeploy.' });
    return out_({ ok: false, error: 'Scan failed (' + msg.slice(0, 120) + ')' });
  }
}

/** Recipe text shared from another app. Nothing is stored in Drive for this. */
function parseText_(text) {
  text = String(text || '').slice(0, 20000);
  if (text.replace(/\s/g, '').length < 20) return out_({ ok: false, error: 'There isn\'t enough text to read as a recipe' });
  return out_({ ok: true, text: text, fileUrl: '', recipe: parseScan_(text) });
}

// ---------- reading text out of files (shared by upload, Drive inbox and pasted Drive links) ----------
/** Text of a Drive file. Google Docs are read directly; everything else goes through textFromBlob_. */
function textFromFile_(file) {
  const mime = file.getMimeType();
  if (mime === 'application/vnd.google-apps.document') return DocumentApp.openById(file.getId()).getBody().getText();
  return textFromBlob_(file.getBlob(), mime, file.getName());
}

function textFromBlob_(blob, mime, name) {
  if (mime === 'text/plain') return blob.getDataAsString('UTF-8');
  if (mime === 'message/rfc822' || /\.eml$/i.test(name || '')) return emlText_(blob.getDataAsString('UTF-8'));
  // Drive turns photos/PDFs (with OCR) and Word files into a Google Doc; we read it, then delete it.
  let id = null;
  try {
    const f = Drive.Files.create({ name: 'recipe-scan (temp)', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'en' });
    id = f.id;
    return DocumentApp.openById(id).getBody().getText();
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* temp file cleanup is best-effort */ } }
  }
}

// ---------- email (.eml): subject becomes the first line (the title), then the body text ----------
function emlText_(raw) {
  raw = String(raw).replace(/\r\n/g, '\n');
  const hdrEnd = raw.indexOf('\n\n');
  const hdr = (hdrEnd < 0 ? raw : raw.slice(0, hdrEnd)).replace(/\n[ \t]+/g, ' ');
  const rawSubj = decodeWords_(((hdr.match(/^Subject:[ \t]*(.*)$/mi) || [])[1] || '')).trim();
  const parts = mimePart_(raw);
  return composeEmail_(rawSubj, parts.plain || htmlToText_(parts.html || ''));
}

/** Subject + body of an email -> recipe text: title first, forwarding clutter removed. */
function composeEmail_(rawSubj, body) {
  rawSubj = String(rawSubj || '').trim();
  let subj = rawSubj.replace(/^\s*((fwd?|fw|re):\s*)+/i, '').trim();
  body = String(body || '').replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim();
  const fw = stripForwarded_(body);
  body = fw.body;
  if (fw.subject && (!subj || /^\s*(fwd?|fw):/i.test(rawSubj))) subj = fw.subject;   // the original email's subject is the better title
  const firstLine = body.split('\n')[0].trim().toLowerCase();
  if (subj && firstLine === subj.toLowerCase()) body = body.split('\n').slice(1).join('\n').trim();   // body repeats the subject
  return (subj ? subj + '\n' : '') + body;
}

/**
 * A forwarded email looks like: [forwarder's note] / "---------- Forwarded message ---------" / From: Date: Subject: To: / [original].
 * Returns the original's body (header block removed), its subject, and keeps a long forwarder note at the end.
 */
function stripForwarded_(body) {
  const lines = String(body).split('\n');
  const mi = lines.findIndex(function (l) { return /^\W*(-{2,}\s*(forwarded message|original message)\s*-{2,}|begin forwarded message:?)\W*$/i.test(l.trim()); });
  if (mi < 0) return { body: body, subject: '' };
  let j = mi + 1, subject = '';
  while (j < lines.length && j < mi + 16) {
    const l = lines[j].trim();
    const hm = l.match(/^\W*(From|Date|Sent|To|Cc|Bcc|Reply-To|Subject)\W*:\s*(.*)$/i);
    if (hm) { if (/^subject$/i.test(hm[1])) subject = hm[2].replace(/^\s*((fwd?|fw|re):\s*)+/i, '').trim(); j++; }
    else if (!l) j++;
    else break;
  }
  const note = lines.slice(0, mi).join('\n').trim();
  const rest = lines.slice(j).join('\n').trim();
  return { body: note.length > 200 ? rest + '\n\n' + note : rest, subject: subject };
}

/** Walks a MIME message and returns the first text/plain and text/html bodies, decoded. */
function mimePart_(raw) {
  const i = raw.indexOf('\n\n');
  const hdr = (i < 0 ? raw : raw.slice(0, i)).replace(/\n[ \t]+/g, ' ');
  const body = i < 0 ? '' : raw.slice(i + 2);
  const h = function (name) { const m = hdr.match(new RegExp('^' + name + ':[ \\t]*(.*)$', 'mi')); return m ? m[1] : ''; };
  const ctype = h('Content-Type').toLowerCase() || 'text/plain';
  const out = { plain: '', html: '' };
  const bm = h('Content-Type').match(/boundary="?([^";]+)"?/i);
  if (/^multipart\//.test(ctype) && bm) {
    body.split('--' + bm[1]).slice(1).forEach(function (seg) {
      if (/^--/.test(seg)) return;                       // closing boundary
      const p = mimePart_(seg.replace(/^\n/, ''));
      if (!out.plain && p.plain) out.plain = p.plain;
      if (!out.html && p.html) out.html = p.html;
    });
    return out;
  }
  if (/attachment/i.test(h('Content-Disposition'))) return out;
  const cs = ((h('Content-Type').match(/charset="?([^";\s]+)"?/i) || [])[1] || 'UTF-8');
  const enc = h('Content-Transfer-Encoding').toLowerCase().trim();
  let text;
  if (enc === 'base64') text = bytesToString_(Utilities.base64Decode(body.replace(/\s+/g, '')), cs);
  else if (enc === 'quoted-printable') text = qpDecode_(body, cs);
  else text = body;
  if (/^text\/html/.test(ctype)) out.html = text; else if (/^text\/plain/.test(ctype)) out.plain = text;
  return out;
}

function bytesToString_(bytes, charset) {
  try { return Utilities.newBlob(bytes).getDataAsString(charset || 'UTF-8'); } catch (e) { return Utilities.newBlob(bytes).getDataAsString('UTF-8'); }
}

function qpDecode_(s, charset) {
  s = s.replace(/=\n/g, '');
  const bytes = [];
  for (let i = 0; i < s.length; i++) {
    const c = s.charAt(i);
    if (c === '=' && /^[0-9A-Fa-f]{2}$/.test(s.substr(i + 1, 2))) { bytes.push(parseInt(s.substr(i + 1, 2), 16)); i += 2; }
    else if (s.charCodeAt(i) < 128) bytes.push(s.charCodeAt(i));
    else Utilities.newBlob(c).getBytes().forEach(function (b) { bytes.push(b); });
  }
  return bytesToString_(bytes.map(function (b) { return b > 127 ? b - 256 : b; }), charset);
}

/** RFC 2047 encoded words in headers: =?UTF-8?B?...?= and =?UTF-8?Q?...?= */
function decodeWords_(s) {
  return String(s).replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, function (_, cs, enc, txt) {
    try {
      if (enc.toUpperCase() === 'B') return bytesToString_(Utilities.base64Decode(txt), cs);
      return qpDecode_(txt.replace(/_/g, ' '), cs);
    } catch (e) { return txt; }
  });
}

function htmlToText_(html) {
  let t = String(html).replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
    .replace(/<li[^>]*>/gi, '\n• ').replace(/<\/(p|div|li|tr|h[1-6]|ul|ol|table)>/gi, '\n').replace(/<br\s*\/?>/gi, '\n');
  return clean_(t).split('\n').map(function (l) { return l.trim(); }).join('\n');
}

/** Best-effort split of OCR text into title / summary / servings / ingredients / steps. */
function parseScan_(text) {
  const raw = String(text).replace(/\r/g, '').split('\n').map(function (l) { return l.replace(/[ \t]+/g, ' ').trim(); });
  const lines = raw.filter(String);
  const isIng = function (l) { return /^(ingredients?)\s*:?$/i.test(l); };
  const isDir = function (l) { return /^(directions?|instructions?|method|steps?|preparation|procedure)\s*:?$/i.test(l); };
  // Nutrition block (usually last): peel it off first so it doesn't leak into the steps.
  const nutI = lines.findIndex(function (l) { return /^nutrition(al)?( information| facts| info)?\s*:?$/i.test(l); });
  const nutrition = nutI >= 0 ? lines.slice(nutI + 1).join('\n') : '';
  if (nutI >= 0) lines.length = nutI;
  const timeOf = function (re) { const m = text.match(re); return m ? m[1].trim().slice(0, 30) : ''; };
  const prepTime = timeOf(/prep(?:aration)?\s*time\s*:?\s*([^\n]+)/i), totalTime = timeOf(/total\s*time\s*:?\s*([^\n]+)/i);
  const ingI = lines.findIndex(isIng), dirI = lines.findIndex(isDir);
  const startsQty = function (l) { return /^[•\-*·]?\s*(\d|[¼½¾⅓⅔⅛])/.test(l); };
  const bullet = function (l) { return l.replace(/^[•\-*·]\s*/, ''); };

  let title = lines[0] || '';
  if (isIng(title) || isDir(title)) title = '';
  let servings = 4;
  const sv = text.match(/(?:serves|servings?|yields?|makes)\s*:?\s*(\d+)/i);
  if (sv) servings = Number(sv[1]);

  let ing = [], stepsRaw = [], summary = '';
  if (ingI >= 0 || dirI >= 0) {
    const ingEnd = (dirI > ingI) ? dirI : lines.length;
    if (ingI >= 0) ing = lines.slice(ingI + 1, ingEnd);
    if (dirI >= 0) stepsRaw = lines.slice(dirI + 1);
    else stepsRaw = [];
    const head = lines.slice(1, ingI >= 0 ? ingI : dirI);
    summary = head.filter(function (l) { return !/^(serves|servings?|yields?|makes|prep|cook|total)/i.test(l); }).join(' ').slice(0, 300);
    if (ingI < 0) { // only a directions heading: ingredient-looking lines before it
      ing = lines.slice(1, dirI).filter(startsQty); summary = '';
    }
  } else { // no headings: quantity-looking lines are ingredients, the rest after them are steps
    const first = lines.findIndex(function (l, i) { return i > 0 && startsQty(l); });
    if (first > 0) {
      let end = first; while (end < lines.length && (startsQty(lines[end]) || /^[a-z(]/.test(lines[end]))) end++;
      ing = lines.slice(first, end); stepsRaw = lines.slice(end);
    } else { stepsRaw = lines.slice(1); }
  }

  // Ingredients: re-join wrapped lines (a line starting lowercase belongs to the one above).
  const ings = [];
  ing.forEach(function (l) {
    l = bullet(l);
    if (!l) return;
    if (ings.length && /^[a-z(]/.test(l) && !startsQty(l)) ings[ings.length - 1] += ' ' + l; else ings.push(l);
  });
  // Steps: numbered markers start new steps; otherwise join wrapped lines until a sentence ends.
  const steps = [];
  stepsRaw.forEach(function (l) {
    const numbered = /^\d{1,2}[.)]\s*\S/.test(l);
    const body = l.replace(/^\d{1,2}[.)]\s*/, '');
    const prevDone = !steps.length || /[.!?)"”]$/.test(steps[steps.length - 1]);
    if (numbered || prevDone) steps.push(body); else steps[steps.length - 1] += ' ' + body;
  });
  return { title: title.slice(0, 120), summary: summary, servings: servings, category: guessCategory_(title + ' ' + summary),
           prepTime: prepTime, totalTime: totalTime, nutrition: nutrition, ingredients: ings, steps: steps };
}

function extract_(html, url) {
  const host = (String(url).match(/^https?:\/\/(?:www\.)?([^\/?#]+)/i) || [])[1] || '';
  const og = function (p) {
    const a = html.match(new RegExp('<meta[^>]+(?:property|name)=["\']' + p + '["\'][^>]*content=["\']([^"\']*)["\']', 'i'));
    const b = html.match(new RegExp('<meta[^>]+content=["\']([^"\']*)["\'][^>]*(?:property|name)=["\']' + p + '["\']', 'i'));
    return clean_((a || b || [])[1] || '');
  };
  const site = og('og:site_name') || host;
  let node = null, m;
  const re = /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  while (!node && (m = re.exec(html))) {
    try { node = findRecipe_(JSON.parse(m[1].trim())); } catch (e) { /* skip malformed block */ }
  }
  if (!node) {
    return { found: false, recipe: { title: og('og:title'), summary: og('og:description'), image: og('og:image'),
      sourceName: site, sourceUrl: url, ingredients: [], steps: [] } };
  }
  const first = function (v) { return Array.isArray(v) ? v[0] : v; };
  let img = first(node.image); if (img && typeof img === 'object') img = img.url || '';
  let author = first(node.author); if (author && typeof author === 'object') author = author.name || '';
  const yieldNum = (String(first(node.recipeYield) || '').match(/\d+/) || [])[0];
  const cat = [].concat(node.recipeCategory || [], node.keywords || []).join(' ') + ' ' + (node.name || '');
  return { found: true, recipe: {
    title: clean_(node.name), summary: clean_(node.description), image: String(img || og('og:image') || ''),
    sourceName: site, sourceUrl: url, author: clean_(author || ''),
    servings: yieldNum ? Number(yieldNum) : 4, category: guessCategory_(cat),
    prepTime: dur_(node.prepTime), totalTime: dur_(node.totalTime), nutrition: nut_(node.nutrition),
    ingredients: [].concat(node.recipeIngredient || node.ingredients || []).map(clean_).filter(String),
    steps: steps_(node.recipeInstructions)
  } };
}

/** ISO-8601 duration ("PT1H15M") -> "1 hr 15 min". Anything else is passed through. */
function dur_(v) {
  if (!v) return '';
  const m = String(v).match(/^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?)?$/i);
  if (!m) return clean_(v);
  const mins = (Number(m[1] || 0) * 24 + Number(m[2] || 0)) * 60 + Number(m[3] || 0);
  if (!mins) return '';
  return mins >= 60 ? Math.floor(mins / 60) + ' hr' + (mins % 60 ? ' ' + (mins % 60) + ' min' : '') : mins + ' min';
}

/** schema.org NutritionInformation -> "Label: value" lines. */
function nut_(n) {
  if (!n || typeof n !== 'object') return '';
  const labels = { servingSize: 'Serving size', calories: 'Calories', fatContent: 'Fat', saturatedFatContent: 'Saturated fat',
    transFatContent: 'Trans fat', cholesterolContent: 'Cholesterol', sodiumContent: 'Sodium', carbohydrateContent: 'Carbohydrates',
    fiberContent: 'Fiber', sugarContent: 'Sugar', proteinContent: 'Protein' };
  return Object.keys(labels).filter(function (k) { return n[k]; }).map(function (k) { return labels[k] + ': ' + clean_(n[k]); }).join('\n');
}

function findRecipe_(x) {
  if (!x || typeof x !== 'object') return null;
  if (Array.isArray(x)) { for (let i = 0; i < x.length; i++) { const r = findRecipe_(x[i]); if (r) return r; } return null; }
  const t = [].concat(x['@type'] || []);
  if (t.indexOf('Recipe') >= 0) return x;
  return findRecipe_(x['@graph']) || findRecipe_(x.mainEntity) || findRecipe_(x.mainEntityOfPage);
}

// Some sites ship every step as one blob: "1. Do this.2. Do that." — split it, and drop the
// leading numbers (the app numbers the list itself).
function splitSteps_(text) {
  let t = String(text).replace(/<\/?(li|p|br)[^>]*>/gi, '\n');
  t = clean_(t);
  if (/^\s*1[.)]\s*\S/.test(t)) t = t.replace(/(^|[.!?)"”])\s*(\d{1,2})[.)]\s*(?=[A-Z])/g, '$1\n');
  return t.split(/\n+/).map(function (s) { return s.replace(/^\s*\d{1,2}[.)]\s+/, '').trim(); }).filter(String);
}

function steps_(ins) {
  if (!ins) return [];
  if (typeof ins === 'string') return splitSteps_(ins);
  const out = [];
  [].concat(ins).forEach(function (s) {
    if (typeof s === 'string') splitSteps_(s).forEach(function (t) { out.push(t); });
    else if (s && s.itemListElement) steps_(s.itemListElement).forEach(function (t) { out.push(t); });
    else if (s && (s.text || s.name)) splitSteps_(s.text || s.name).forEach(function (t) { out.push(t); });
  });
  return out;
}

function guessCategory_(s) {
  s = String(s).toLowerCase();
  if (/breakfast|brunch|pancake|waffle|omelet|oatmeal/.test(s)) return 'Breakfast';
  if (/dessert|cookie|\bcakes?\b|cheesecake|cupcake|pie\b|brownie|ice cream|cobbler/.test(s)) return 'Dessert';
  if (/appetizer|snack|dip\b|starter/.test(s)) return 'Appetizer';
  if (/drink|beverage|cocktail|smoothie|lemonade/.test(s)) return 'Drink';
  if (/side dish|\bside\b|salad|slaw/.test(s)) return 'Side';
  return 'Entrée';
}

function clean_(s) {
  return String(s == null ? '' : s).replace(/<[^>]*>/g, ' ')
    .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(Number(d)); })
    .replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&').replace(/[ \t]+/g, ' ').trim();
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // Keep the header row in step with COLS (adds new columns to an existing sheet).
  const hdr = sh.getRange(1, 1, 1, COLS.length);
  if (hdr.getValues()[0].join('|') !== COLS.join('|')) hdr.setNumberFormat('@').setValues([COLS]).setFontWeight('bold');
  // Plain-text format so nothing a recipe contains is ever treated as a formula.
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), COLS.length).setNumberFormat('@');
  return sh;
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, COLS.length).getValues().map(function (row) {
    const o = {};
    COLS.forEach(function (c, i) { o[c] = row[i]; });
    return {
      id: String(o.id), title: o.title, category: o.category, sourceName: o.sourceName, sourceUrl: o.sourceUrl,
      servings: Number(o.servings) || 1, summary: o.summary, image: o.image,
      ingredients: parse_(o.ingredients), steps: parse_(o.steps), addedBy: o.addedBy, added: o.added,
      favorite: String(o.favorite) === 'true', prepTime: o.prepTime, totalTime: o.totalTime, nutrition: o.nutrition, notes: o.notes
    };
  }).filter(function (r) { return r.id && r.title; }).reverse();
}

function parse_(s) { try { return JSON.parse(s) || []; } catch (e) { return []; } }

function save_(r, name) {
  if (!r || typeof r.title !== 'string' || !r.title.trim() || !r.id) return out_({ ok: false, error: 'invalid recipe' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const lim = function (v, n) { return String(v == null ? '' : v).slice(0, n); };
    const row = [
      lim(r.id, 40), lim(r.title, 200), lim(r.category, 40), lim(r.sourceName, 200), lim(r.sourceUrl, 1000),
      String(Number(r.servings) || 1), lim(r.summary, 2000), lim(r.image, 2000),
      lim(JSON.stringify(r.ingredients || []), 20000), lim(JSON.stringify(r.steps || []), 20000),
      lim(name, 60), new Date().toISOString(),
      String(r.favorite === true), lim(r.prepTime, 40), lim(r.totalTime, 40), lim(r.nutrition, 2000), lim(r.notes, 4000)
    ];
    const n = sh.getLastRow() - 1;
    let rowNum = 0;
    if (n > 0) {
      const ids = sh.getRange(2, 1, n, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === row[0]) { rowNum = i + 2; break; }
    }
    if (rowNum) {
      // keep original "added by" and date when editing
      const old = sh.getRange(rowNum, 11, 1, 2).getValues()[0];
      row[10] = old[0] || row[10]; row[11] = old[1] || row[11];
    } else {
      rowNum = sh.getLastRow() + 1;
    }
    sh.getRange(rowNum, 1, 1, COLS.length).setNumberFormat('@').setValues([row]);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
