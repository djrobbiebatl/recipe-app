/**
 * Recipe API — paste into Extensions > Apps Script of the recipes Google Sheet.
 * Then: Project Settings > Script properties > add PASSCODE = your shared passcode,
 * and Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The passcode is the only gate, so make it a real one (12+ characters).
 */
const SHEET_NAME = 'Recipes';
const COLS = ['id', 'title', 'category', 'sourceName', 'sourceUrl', 'servings', 'summary',
              'image', 'ingredients', 'steps', 'addedBy', 'added', 'favorite', 'prepTime', 'totalTime', 'nutrition', 'notes', 'cookTime'];

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
  if (req.action === 'shoplist') return out_({ ok: true, items: shopRows_().map(shopOut_) });
  if (req.action === 'shopadd') return shopAdd_(req.items, req.recipe);
  if (req.action === 'shopset') return shopSet_(req.id, req.done);
  if (req.action === 'shopdelete') return shopDelete_(req.id, req.mode);
  if (req.action === 'planlist') return out_({ ok: true, plan: planRows_().map(planOut_) });
  if (req.action === 'planadd') return planAdd_(req.date, req.recipeId, req.title);
  if (req.action === 'plandelete') return planDelete_(req.id);
  if (req.action === 'websuggest') return webSuggest_(req.terms);
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
  const id = String(item.id).slice(0, 80);
  // Cheap checks first, so a repeat or a full queue doesn't cost any scanning time.
  const before = intakeRows_();
  if (before.some(function (r) { return r.id === id; })) return out_({ ok: true, duplicate: true });
  if (before.filter(function (r) { return r.status === 'pending'; }).length >= 150) return out_({ ok: false, error: 'queue is full' });

  // Attachments (screenshots, photos, PDFs, Word files) are read the same way as an uploaded file: Drive OCR.
  const okMime = /^(image\/(jpeg|png|webp|gif)|application\/pdf|application\/msword|application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document|text\/plain)$/;
  const scanned = [];
  (item.files || []).slice(0, 3).forEach(function (f) {
    if (!f || !f.data || !okMime.test(String(f.mime))) return;
    try {
      const t = textFromBlob_(Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, String(f.name || 'attachment').slice(0, 100)), f.mime, f.name);
      if (t && t.replace(/\s/g, '').length >= 20) scanned.push(t.trim());
    } catch (e) { /* one unreadable attachment shouldn't lose the email */ }
  });
  const ocr = scanned.join('\n\n');
  const bodyText = String(item.plain || '').slice(0, 60000);
  const bodyIsThin = bodyText.replace(/\s/g, '').length < 40;
  // A screenshot-only email: the recipe's own title (first line of the scan) beats a subject like "Fwd: look at this".
  const combined = (ocr && bodyIsThin) ? ocr : composeEmail_(String(item.subject || '').slice(0, 300), bodyText + (ocr ? '\n\n' + ocr : ''));

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const rows = intakeRows_();
    if (rows.some(function (r) { return r.id === id; })) return out_({ ok: true, duplicate: true });
    const text = combined.slice(0, 40000);
    const empty = text.replace(/\s/g, '').length < 20;
    const row = [id, String(item.from || '').slice(0, 200), text.split('\n')[0].slice(0, 200), String(item.date || new Date().toISOString()).slice(0, 40),
                 text, empty ? 'skipped' : 'pending', new Date().toISOString()];
    const sh = intakeSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, 1, INTAKE_COLS.length).setNumberFormat('@').setValues([row]);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

/**
 * Queued text is re-cleaned every time it's read, so improvements to the cleaning rules also fix emails that were
 * already waiting in the queue (the stored text is the first line = subject, then the body).
 */
function intakeClean_(stored) {
  const nl = String(stored).indexOf('\n');
  if (nl < 0) return String(stored);
  return composeEmail_(stored.slice(0, nl), stored.slice(nl + 1));
}

function intakeListAction_() {
  const items = intakeRows_().filter(function (r) { return r.status === 'pending'; }).reverse().map(function (r) {
    const t = intakeClean_(r.text);
    return { id: r.id, from: r.from, subject: t.split('\n')[0] || r.subject, date: r.date, snippet: t.split('\n').slice(1).join(' ').replace(/\s+/g, ' ').slice(0, 180) };
  });
  return out_({ ok: true, address: PropertiesService.getScriptProperties().getProperty('INTAKE_ADDRESS') || '', items: items });
}

function intakeCountAction_() {
  return out_({ ok: true, count: intakeRows_().filter(function (r) { return r.status === 'pending'; }).length });
}

function intakeGetAction_(id) {
  const r = intakeRows_().filter(function (x) { return x.id === String(id); })[0];
  if (!r) return out_({ ok: false, error: 'That email isn\'t in the queue any more' });
  const t = intakeClean_(r.text);
  return out_({ ok: true, text: t.slice(0, 20000), fileUrl: '', from: r.from, recipe: parseScan_(t) });
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

// ---------- weekly shopping list: shared by everyone, kept in the "Shopping" tab ----------
const SHOP_COLS = ['id', 'item', 'amt', 'unit', 'recipes', 'done', 'added'];

function shopSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Shopping');
  if (!sh) {
    sh = ss.insertSheet('Shopping');
    sh.getRange(1, 1, 1, SHOP_COLS.length).setNumberFormat('@').setValues([SHOP_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function shopRows_() {
  const sh = shopSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, SHOP_COLS.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0]), item: String(r[1]), amt: String(r[2]), unit: String(r[3]), recipes: String(r[4]), done: String(r[5]) === 'yes', added: String(r[6]) };
  });
}

function shopOut_(r) {
  return { id: r.id, item: r.item, amt: r.amt === '' ? null : Number(r.amt), unit: r.unit, recipes: r.recipes ? r.recipes.split(' | ') : [], done: r.done };
}

const SHOP_UNITS = { tablespoon: 'tbsp', tablespoons: 'tbsp', tbsps: 'tbsp', teaspoon: 'tsp', teaspoons: 'tsp', tsps: 'tsp', cups: 'cup', ounce: 'oz', ounces: 'oz',
  pound: 'lb', pounds: 'lb', lbs: 'lb', gram: 'g', grams: 'g', cloves: 'clove', cans: 'can', slices: 'slice', packages: 'package', pkg: 'package', pkgs: 'package',
  packets: 'packet', pkt: 'packet', pkts: 'packet', sticks: 'stick', bunches: 'bunch', heads: 'head', jars: 'jar', bottles: 'bottle', boxes: 'box', sprigs: 'sprig', stalks: 'stalk' };

function shopUnitKey_(u) {
  const k = String(u || '').toLowerCase().replace(/\./g, '').trim();
  return SHOP_UNITS[k] || k;
}
function shopItemKey_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim().replace(/(ies)$/, 'y').replace(/(es)$/, '').replace(/s$/, '');
}

/** items: [{amt, unit, item}] from one recipe (or typed by hand, recipe ''). Same item + same unit adds up; otherwise a new line. */
function shopAdd_(items, recipe) {
  if (!items || !items.length) return out_({ ok: false, error: 'nothing to add' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = shopSheet_();
    const rows = shopRows_();
    const stamp = new Date().toISOString();
    let merged = 0, created = 0;
    items.slice(0, 80).forEach(function (it) {
      const name = String(it.item || '').trim().slice(0, 200);
      if (!name) return;
      const amt = (it.amt === null || it.amt === undefined || it.amt === '' || isNaN(Number(it.amt))) ? null : Number(it.amt);
      const unit = String(it.unit || '').trim().slice(0, 30);
      const hit = rows.filter(function (r) {
        return !r.done && shopItemKey_(r.item) === shopItemKey_(name) && shopUnitKey_(r.unit) === shopUnitKey_(unit) && (r.amt === '') === (amt === null);
      })[0];
      if (hit) {
        const list = hit.recipes ? hit.recipes.split(' | ') : [];
        if (recipe && list.indexOf(recipe) < 0) list.push(recipe);
        hit.recipes = list.join(' | ');
        if (amt !== null) hit.amt = String(Math.round((Number(hit.amt) + amt) * 1000) / 1000);
        sh.getRange(hit.row, 3, 1, 1).setValue(hit.amt);
        sh.getRange(hit.row, 5, 1, 1).setValue(hit.recipes);
        merged++;
      } else {
        const row = [Utilities.getUuid(), name, amt === null ? '' : String(amt), unit, recipe ? String(recipe).slice(0, 120) : '', '', stamp];
        sh.appendRow(row);
        rows.push({ row: sh.getLastRow(), id: row[0], item: name, amt: row[2], unit: unit, recipes: row[4], done: false, added: stamp });
        created++;
      }
    });
    return out_({ ok: true, added: created, merged: merged, items: shopRows_().map(shopOut_) });
  } finally { lock.releaseLock(); }
}

function shopSet_(id, done) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = shopRows_().filter(function (x) { return x.id === String(id); })[0];
    if (r) shopSheet_().getRange(r.row, 6).setValue(done ? 'yes' : '');
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

/** mode 'one' (with id), 'done' (ticked items), or 'all'. */
function shopDelete_(id, mode) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = shopSheet_();
    const rows = shopRows_();
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      if ((mode === 'one' && r.id === String(id)) || (mode === 'done' && r.done) || mode === 'all') sh.deleteRow(r.row);
    }
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

// ---------- dinner calendar: recipes planned for a date, kept in the "Plan" tab (several per night allowed) ----------
const PLAN_COLS = ['id', 'date', 'recipeId', 'title', 'added'];

function planSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Plan');
  if (!sh) {
    sh = ss.insertSheet('Plan');
    sh.getRange(1, 1, 1, PLAN_COLS.length).setNumberFormat('@').setValues([PLAN_COLS]).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('@'); // keep dates as plain 2026-11-20 text
    sh.setFrozenRows(1);
  }
  return sh;
}

function planDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

function planRows_() {
  const sh = planSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, PLAN_COLS.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0]), date: planDate_(r[1]), recipeId: String(r[2]), title: String(r[3]) };
  });
}

function planOut_(r) { return { id: r.id, date: r.date, recipeId: r.recipeId, title: r.title }; }

function planAdd_(date, recipeId, title) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date)) || !recipeId) return out_({ ok: false, error: 'pick a date' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rows = planRows_();
    const dup = rows.filter(function (r) { return r.date === date && r.recipeId === String(recipeId); })[0];
    if (!dup) planSheet_().appendRow([Utilities.getUuid(), date, String(recipeId), String(title || '').slice(0, 200), new Date().toISOString()]);
    return out_({ ok: true, already: !!dup, plan: planRows_().map(planOut_) });
  } finally { lock.releaseLock(); }
}

function planDelete_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = planRows_().filter(function (x) { return x.id === String(id); })[0];
    if (r) planSheet_().deleteRow(r.row);
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}

// ---------- Pantry tab: web suggestions from Spoonacular's free plan (key is the Script property SPOONACULAR_KEY) ----------
const WEB_WEAK = ['mix', 'fresh', 'frozen', 'canned', 'dried', 'boneless', 'skinless', 'whole', 'large', 'small', 'medium', 'of', 'the', 'a'];
function webWords_(s) {
  return String(s).toLowerCase().replace(/[^a-z\s]/g, ' ').split(/\s+/).filter(Boolean).map(function (w) {
    return /ies$/.test(w) ? w.slice(0, -3) + 'y' : /oes$/.test(w) ? w.slice(0, -2) : (/s$/.test(w) && !/ss$/.test(w)) ? w.slice(0, -1) : w;
  });
}

/** Adds this call's cost to acc.request and records points used today / left, from the response headers (any capitalisation). */
function webQuota_(resp, acc) {
  let h = {};
  try { h = resp.getHeaders() || {}; } catch (e) {}
  const get = function (name) {
    for (const k in h) if (String(k).toLowerCase() === name) { const n = parseFloat(h[k]); return isNaN(n) ? null : n; }
    return null;
  };
  const r = get('x-api-quota-request'), u = get('x-api-quota-used'), l = get('x-api-quota-left');
  if (r !== null) acc.request = Math.round((acc.request + r) * 100) / 100;
  if (u !== null) acc.used = u;
  if (l !== null) acc.left = l;
}

function webSuggest_(terms) {
  const key = PropertiesService.getScriptProperties().getProperty('SPOONACULAR_KEY');
  if (!key) return out_({ ok: false, error: 'The web search key isn\'t set up yet (Script property SPOONACULAR_KEY).' });
  const list = (terms || []).map(function (t) { return String(t).trim().slice(0, 60); }).filter(Boolean).slice(0, 8);
  if (!list.length) return out_({ ok: false, error: 'Type a few ingredients first' });
  const opts = { muteHttpExceptions: true, headers: { 'x-api-key': key } };
  const base = 'https://api.spoonacular.com/recipes/';
  // Spoonacular treats "include these ingredients" as "any of these", so ask for the best 100 (most-used first) and keep only
  // recipes that use EVERY ingredient typed.
  const quota = { request: 0, used: null, left: null };   // from Spoonacular's X-API-Quota-* response headers
  const found = UrlFetchApp.fetch(base + 'complexSearch?number=100&sort=max-used-ingredients&fillIngredients=true&includeIngredients=' + encodeURIComponent(list.join(',')), opts);
  const code = found.getResponseCode();
  webQuota_(found, quota);
  if (code === 402 || code === 429) return out_({ ok: false, code: 'quota', error: 'We\'re out of web-search credits for today. They reset at midnight UTC — your saved recipes still work.', quota: quota });
  if (code === 401 || code === 403) return out_({ ok: false, error: 'The web search key was rejected — check SPOONACULAR_KEY.' });
  if (code !== 200) return out_({ ok: false, error: 'The web search service had a problem (' + code + '). Try again later.' });
  const hits = (JSON.parse(found.getContentText()).results || []).filter(function (h) {
    const used = (h.usedIngredients || []).map(function (i) { return webWords_(i.name); });
    return list.every(function (t) {
      const all = webWords_(t), core = all.filter(function (w) { return WEB_WEAK.indexOf(w) < 0; }), need = core.length ? core : all;
      return need.length && used.some(function (u) { return need.every(function (w) { return u.indexOf(w) >= 0; }); });
    });
  }).slice(0, 6);
  if (!hits.length) return out_({ ok: true, results: [], quota: quota });
  // The first call has no web address for the recipe; one bulk lookup adds it (plus time and servings).
  const info = {};
  const bulk = UrlFetchApp.fetch(base + 'informationBulk?ids=' + hits.map(function (h) { return h.id; }).join(','), opts);
  webQuota_(bulk, quota);
  if (bulk.getResponseCode() === 402 || bulk.getResponseCode() === 429) return out_({ ok: false, code: 'quota', error: 'We\'re out of web-search credits for today. They reset at midnight UTC — your saved recipes still work.', quota: quota });
  if (bulk.getResponseCode() === 200) JSON.parse(bulk.getContentText()).forEach(function (x) { info[x.id] = x; });
  const names = function (a) { return (a || []).map(function (i) { return i.name; }); };
  return out_({ ok: true, quota: quota, results: hits.map(function (h) {
    const x = info[h.id] || {};
    return { id: h.id, title: h.title, image: h.image || '', url: x.sourceUrl || '', sourceName: x.sourceName || '', page: x.spoonacularSourceUrl || '',
             minutes: x.readyInMinutes || null, servings: x.servings || null, used: names(h.usedIngredients), missed: names(h.missedIngredients) };
  }) });
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
  url = stripTracking_(url);
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
  // "Fwd:", "Emailing:", "Scanned:" prefixes and a trailing file extension (".doc") are not part of a recipe name.
  let subj = rawSubj.replace(/^\s*((fwd?|fw|re|emailing|scanned|scan|attached|attachment)\s*:\s*)+/i, '').replace(/\.(docx?|pdf|txt|rtf|jpe?g|png)\s*$/i, '').trim();
  body = String(body || '').replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim();
  const fw = stripForwarded_(body);
  body = cleanEmailBody_(fw.body);
  if (fw.subject && (!subj || /^\s*(fwd?|fw):/i.test(rawSubj))) subj = fw.subject;   // the original email's subject is the better title
  // A body line that is the subject minus punctuation ("Shepherds" vs "Shepherd's") is the recipe's real title: use it as written.
  const alnum = function (x) { return String(x).toLowerCase().replace(/[^a-z0-9]/g, ''); };
  if (subj) {
    const bl = body.split('\n');
    const at = bl.slice(0, 40).findIndex(function (l) { return l.trim() && alnum(l) === alnum(subj); });
    if (at >= 0) { subj = bl[at].trim(); bl.splice(at, 1); body = bl.join('\n').trim(); }
  }
  const firstLine = body.split('\n')[0].trim().toLowerCase();
  const s = subj.toLowerCase();
  if (subj && firstLine === s) body = body.split('\n').slice(1).join('\n').trim();   // body repeats the subject
  // Subject "chili", body starts "Aunt Jo's Chili": the body's own line is the real title, so let it lead.
  else if (subj && firstLine && firstLine.length <= 80 && (firstLine.indexOf(s) >= 0 || s.indexOf(firstLine) >= 0)) subj = '';
  return (subj ? subj + '\n' : '') + body;
}

/**
 * Email leftovers that don't belong in a recipe: forwarding header lines (From/Sent/Date/To/Subject) anywhere in the
 * text, "On <date>, <name> wrote:" lines, "Sent from my iPhone" style footers, and ">" quote marks.
 */
function cleanEmailBody_(body) {
  const hdr = /^\W*(From|Sent|Date|To|Cc|Bcc|Reply-To|Subject)\W*:/i;
  const lines = String(body).split('\n').map(function (l) { return l.replace(/^(\s*>)+ ?/, ''); });
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i].trim();
    if (/^[-_=*~]{4,}$/.test(l)) continue;                                              // "-------" rules
    if (/^(emailing|scanned|scan|attached|attachment)\s*:\s*.+\.(docx?|pdf|txt|rtf|jpe?g|png)\s*$/i.test(l)) continue;   // copier/scanner filename line
    if (hdr.test(l)) {
      let j = i;
      while (j < lines.length && hdr.test(lines[j].trim())) j++;
      // two or more in a row is a header block; a lone one only counts if it carries an address or a time
      if (j - i >= 2 || /@|\b\d{1,2}:\d{2}\b|\b(am|pm)\b/i.test(l)) { i = j - 1; continue; }
    }
    if (/^on .{5,160}\bwrote:\s*$/i.test(l) || /^(-+\s*)?sent from my\b/i.test(l) || /^get outlook for\b/i.test(l) || /^sent (with|via|using) /i.test(l)) continue;
    out.push(lines[i]);
  }
  // Legal / "think of the environment" footers are whole paragraphs; drop any paragraph that reads like one.
  const boiler = /\b(intended (exclusively|solely|only) for|named addressee|may contain (information that is )?(confidential|privileged)|confidential and (may be )?privileged|if you (have )?received this (e-?mail|communication|message|transmission) in error|consider the environment|think before you print|please consider (the )?environment|unauthori[sz]ed (use|disclosure|review))\b/i;
  return out.join('\n').replace(/\n{3,}/g, '\n\n').split(/\n{2,}/).filter(function (p) { return !boiler.test(p); }).join('\n\n').trim();
}

/**
 * The Nutrition block is usually last in an email, so whatever follows it (a signature, a name, a company) would be
 * swept in. Keep only lines that read like nutrition facts, and stop at the first line that doesn't.
 */
function nutLines_(arr) {
  const word = /\b(calories?|kcal|energy|fat|saturated|trans|cholesterol|sodium|salt|carbohydrates?|carbs?|fib(?:er|re)|sugars?|protein|vitamin|calcium|iron|potassium|serving|servings|daily value|amount)\b/i;
  const value = /^[\d.,]+(\s*[-–]\s*[\d.,]+)?\s*(g|mg|mcg|kcal|cal|oz|%)?\s*(\d+\s*%)?$/i;   // "120", "3 g", "28-32g", "11%"
  const kept = [];
  for (let i = 0; i < arr.length; i++) {
    const l = arr[i].trim().replace(/^[•·]\s*/, '');
    if (!l) continue;
    if ((word.test(l) && l.split(/\s+/).length <= 8 && !/[.!?]$/.test(l)) || value.test(l)) kept.push(l);   // a sentence about nutrition isn't a nutrition fact
    else if (!kept.length && !/\d/.test(l) && l.split(/\s+/).length <= 4) continue;   // a lead-in such as "Per serving:" or "Nutrition Component"
    else break;
  }
  // Tables come through as a label line then a value line: "Calories" / "320-380" -> "Calories: 320-380".
  const out = [];
  for (let i = 0; i < kept.length; i++) {
    if (!value.test(kept[i]) && !/\d/.test(kept[i]) && i + 1 < kept.length && value.test(kept[i + 1])) { out.push(kept[i] + ': ' + kept[i + 1]); i++; }
    else out.push(kept[i]);
  }
  return out;
}

/**
 * Web-article lists: sub-headings ("Honey Garlic Sauce", "Prep the chicken") sit between bulleted items. When a section is
 * mostly bullets, drop the short capitalised non-bullet lines (the sub-headings) and strip the bullet marks.
 */
function listLines_(arr) {
  const bullet = /^[•·]\s*/;
  if (arr.filter(function (l) { return bullet.test(l); }).length < 2) return arr;
  return arr.filter(function (l) {
    return bullet.test(l) || !(/^[A-Z0-9(]/.test(l) && l.split(/\s+/).length <= 8 && !/[.!?]$/.test(l));
  }).map(function (l) { return l.replace(bullet, ''); });
}

/** Some recipe blogs list each ingredient as: Name / "Quantity: 2 tbsp olive oil" / "Notes: ...". Returns the plain ingredient lines, or null if not this style. */
function qtyFormat_(arr) {
  if (!arr.some(function (l) { return /^quantity\s*:/i.test(l); })) return null;
  const out = [];
  let name = '', used = true;
  const loose = /^(to taste|as needed|optional|for serving|for garnish|desired|a pinch|pinch|dash)\b/i;
  arr.forEach(function (l) {
    const q = l.match(/^quantity\s*:\s*(.*)$/i);
    if (q) {
      const t = q[1].replace(/[.;]\s*$/, '').trim();
      out.push(loose.test(t) && name ? name + ', ' + t : (t || name));   // "Salt & pepper" + "to taste"
      used = true;
    } else if (/^notes?\s*:/i.test(l)) { /* notes belong to the recipe author, not the ingredient line */ }
    else { if (name && !used) out.push(name); name = l; used = false; }
  });
  if (name && !used) out.push(name);
  return out.filter(String);
}

/** Drops a sign-off and everything after it ("Thanks, Jo", phone, address, "Sent from my iPhone") from the end of scanned lines. */
function trimSignature_(lines) {
  // A sign-off word, optionally followed by a name (up to two capitalised words): "Thanks," / "Best, Jo". "Best served warm." is not one.
  const closing = { test: function (l) {
    const m = l.match(/^(thanks|thank you|best|best regards|kind regards|warm regards|regards|cheers|sincerely|yours truly|talk soon|xoxo)\b/i);
    if (!m) return false;
    return /^([A-Z][\w'-]*\s*){0,2}$/.test(l.slice(m[0].length).replace(/^[,!.]\s*/, '').trim());
  } };
  const junk = /^(--+|_{3,})$|^sent from my\b|^get outlook\b|@[\w.-]+\.\w+|^\+?[\d\s().-]{10,}$|^https?:\/\/\S+$|^www\.\S+$/i;
  const from = Math.max(0, lines.length - 10);
  for (let i = lines.length - 1; i >= from; i--) {
    if (closing.test(lines[i].trim())) { lines.length = i; break; }
  }
  // No sign-off word? A signature still gives itself away with an email address or phone number near the end.
  // Cut from there, taking the short name / title / company lines just above it (never lines with digits: those are nutrition or amounts).
  const contact = /@[\w.-]+\.\w+|\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}/;
  const tail = Math.max(0, lines.length - 8);
  let c = -1;
  for (let i = tail; i < lines.length; i++) { if (contact.test(lines[i])) { c = i; break; } }
  if (c >= 0) {
    let s = c;
    while (s > tail && !/\d/.test(lines[s - 1]) && !/[.!?:]$/.test(lines[s - 1].trim()) && lines[s - 1].trim().split(/\s+/).length <= 5) s--;
    lines.length = s;
  }
  while (lines.length && junk.test(lines[lines.length - 1].trim())) lines.length--;
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
  trimSignature_(lines);
  // Headings may carry a little extra: "Ingredients (Serves 4)", "Instructions:", "Cooking Directions". Not sentences, though.
  const isIng = function (l) { return /^ingredients?\b[^.!?]{0,40}$/i.test(l) && !/^\W/.test(l); };
  const isDir = function (l) { return /^((cooking|baking) )?(directions?|instructions?|method|steps?|preparation|procedure)\b[^.!?]{0,40}$/i.test(l) && !/^\W/.test(l); };
  // Nutrition block (usually last): peel it off first so it doesn't leak into the steps.
  const nutI = lines.findIndex(function (l) { return /^nutrition(al)?( information| facts| info| value)?\b[^.!?]{0,40}$/i.test(l) && !/^\W/.test(l); });
  const nutrition = nutI >= 0 ? nutLines_(lines.slice(nutI + 1)).join('\n') : '';
  if (nutI >= 0) lines.length = nutI;
  const timeOf = function (re) { const m = text.match(re); return m ? m[1].trim().slice(0, 30) : ''; };
  const cookTime = timeOf(/cook(?:ing)?\s*time\s*:?\s*([^\n]+)/i);
  const prepTime = timeOf(/prep(?:aration)?\s*time\s*:?\s*([^\n]+)/i), totalTime = timeOf(/total\s*time\s*:?\s*([^\n]+)/i);
  const ingI = lines.findIndex(isIng), dirI = lines.findIndex(isDir);
  const startsQty = function (l) { return /^[•\-*·]?\s*(\d|[¼½¾⅓⅔⅛])/.test(l); };
  const bullet = function (l) { return l.replace(/^[•\-*·]\s*/, ''); };

  let title = lines[0] || '';
  if (isIng(title) || isDir(title)) title = '';
  let servings = 4;
  const sv = text.match(/(?:serves|servings?|yields?|makes)\s*:?\s*(\d+)/i);
  if (sv) servings = Number(sv[1]);

  let ing = [], stepsRaw = [], summary = '', narIngs = null, narSteps = null;
  if (ingI >= 0 || dirI >= 0) {
    const ingEnd = (dirI > ingI) ? dirI : lines.length;
    if (ingI >= 0) ing = listLines_(lines.slice(ingI + 1, ingEnd));
    if (ingI >= 0) { const q = qtyFormat_(ing); if (q) ing = q; }
    if (dirI >= 0) stepsRaw = listLines_(lines.slice(dirI + 1));
    else stepsRaw = [];
    const head = lines.slice(1, ingI >= 0 ? ingI : dirI);
    summary = head.filter(function (l) { return !/^(serves|servings?|yields?|makes|prep|cook|total)/i.test(l); }).join(' ').slice(0, 300);
    if (ingI < 0) { // only a directions heading: ingredient-looking lines before it
      ing = lines.slice(1, dirI).filter(startsQty); summary = '';
    }
  } else { // no headings: quantity-looking lines are ingredients, the rest after them are steps
    const paras = paragraphs_(raw, lines.length);
    if (paras.length) { paras[0].shift(); if (!paras[0].length) paras.shift(); }   // the title line
    const nar = narrative_(paras);
    const first = lines.findIndex(function (l, i) { return i > 0 && startsQty(l); });
    if (nar) {
      narIngs = nar.ings; narSteps = nar.steps;
    } else if (first > 0) {
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
  if (narIngs) { ings.length = 0; narIngs.forEach(function (x) { ings.push(x); }); }
  // "11/2 lbs" is how some sites write "1 1/2 lbs": a digit glued to a proper fraction.
  for (let i = 0; i < ings.length; i++) {
    ings[i] = ings[i].replace(/^([1-9])([1-9])\/([2-9])\b/, function (m, a, b, c) { return Number(b) < Number(c) ? a + ' ' + b + '/' + c : m; });
  }
  // Steps: numbered markers start new steps; otherwise join wrapped lines until a sentence ends.
  const steps = narSteps || [];
  if (!narSteps) stepsRaw.forEach(function (l) {
    const numbered = /^\d{1,2}[.)]\s*\S/.test(l);
    const body = l.replace(/^\d{1,2}[.)]\s*/, '');
    const prevDone = !steps.length || /[.!?)"”]$/.test(steps[steps.length - 1]);
    if (numbered || prevDone) steps.push(body); else steps[steps.length - 1] += ' ' + body;
  });
  return { title: title.slice(0, 120), summary: summary, servings: servings, category: guessCategory_(title + ' ' + summary),
           prepTime: prepTime, cookTime: cookTime, totalTime: totalTime, nutrition: nutrition, ingredients: ings, steps: steps };
}

/** Non-blank lines grouped into paragraphs (split at blank lines), keeping only the first `count` non-blank lines. */
function paragraphs_(raw, count) {
  const paras = []; let cur = [], n = 0;
  for (let i = 0; i < raw.length; i++) {
    if (!raw[i]) { if (cur.length) { paras.push(cur); cur = []; } continue; }
    if (n >= count) break;
    cur.push(raw[i]); n++;
  }
  if (cur.length) paras.push(cur);
  return paras;
}

/**
 * Recipes written as prose with no "Ingredients" heading, e.g.
 *   Brown 1lb ground beef with / Salt / Pepper / Minced garlic        (lead-in line, then a short list)
 *   Place back over high heat and add / 2 pkt. Gravy mix / 2 cups water / One cup frozen veggies
 * Finds the lists inside each paragraph and pulls them out as ingredients; the text around them stays as steps
 * (the lead-in keeps its list inline so the step still reads naturally). Returns null if no list is found.
 */
function narrative_(paras) {
  const NUM = { one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10 };
  const words = function (l) { return l.split(/\s+/).length; };
  const qty = function (l) {
    return /^[•\-*·]?\s*(\d|[¼½¾⅓⅔⅛]|(one|two|three|four|five|six|seven|eight|nine|ten)\b)/i.test(l) &&
      !/^\d{1,2}[.)]\s/.test(l) && (!/[.!?]$/.test(l) || words(l) <= 6);
  };
  const short = function (l) { return words(l) <= 12 && !/[.!?]$/.test(l); };
  const leadIn = function (l) { return /(\b(with|add|and add|plus|including|combine|mix in|stir in)|:)\s*$/i.test(l); };
  const norm = function (l) {
    return l.replace(/^[•\-*·]\s*/, '').replace(/^(one|two|three|four|five|six|seven|eight|nine|ten)\b/i, function (m) { return NUM[m.toLowerCase()]; })
      .replace(/^(\d+(?:\.\d+)?)(lbs?|oz|g|kg|ml|cups?|tbsp|tsp)\b/i, '$1 $2').trim();
  };
  const ings = [], steps = [];
  let found = false;
  const pushSteps = function (ls) {          // join wrapped lines until a sentence ends
    const out = [];
    ls.forEach(function (l) {
      const numbered = /^\d{1,2}[.)]\s*\S/.test(l), body = l.replace(/^\d{1,2}[.)]\s*/, '');
      if (numbered || !out.length || /[.!?)"”]$/.test(out[out.length - 1])) out.push(body); else out[out.length - 1] += ' ' + body;
    });
    out.forEach(function (s) { steps.push(s); });
  };
  paras.forEach(function (p) {
    let k = -1;
    for (let i = 0; i < p.length; i++) {
      if (qty(p[i]) || (i > 0 && leadIn(p[i - 1]) && short(p[i]))) { k = i; break; }
    }
    if (k < 0) { pushSteps(p); return; }
    let j = k;
    while (j < p.length && (qty(p[j]) || short(p[j]))) j++;
    const items = p.slice(k, j).map(norm).filter(String);
    const lead = p.slice(0, k);
    found = true;
    if (lead.length) {
      const leadText = lead.join(' ');
      const q = leadText.match(/(\d+(?:\.\d+)?\s?(?:lbs?|oz|g|kg|ml|cups?|tbsp|tsp|pkt|pkg|cans?|cloves?)\b\.?\s+[a-z][a-z' -]*?)(?=\s+(?:with|and|to|in|over|until)\b|$)/i);
      if (q) ings.push(norm(q[1]));
      steps.push(leadIn(lead[lead.length - 1]) ? leadText + ' ' + items.join(', ') + '.' : leadText);
    }
    items.forEach(function (x) { ings.push(x); });
    if (j < p.length) pushSteps(p.slice(j));
  });
  return found ? { ings: ings, steps: steps } : null;
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
    // No embedded recipe card: many blogs are just an article with an "Ingredients" list and "Instructions" steps. Read it as text.
    const art = articleRecipe_(html, url, og, site);
    if (art) return { found: true, recipe: art };
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
    servings: yieldNum ? Number(yieldNum) : 4, category: guessCategory_(cat) + extraTags_(cat),
    prepTime: dur_(node.prepTime), cookTime: dur_(node.cookTime), totalTime: dur_(node.totalTime), nutrition: nut_(node.nutrition),
    ingredients: [].concat(node.recipeIngredient || node.ingredients || []).map(clean_).filter(String),
    steps: steps_(node.recipeInstructions)
  } };
}

/** ", Crock Pot" / ", Instant Pot" when the title or keywords say so (recipes can carry several categories). */
function extraTags_(s) {
  return (/slow[- ]?cooker|crock[- ]?pot/i.test(s) ? ', Crock Pot' : '') + (/instant[- ]?pot|pressure cooker/i.test(s) ? ', Instant Pot' : '');
}

/** The part of a page that holds the article: marked up as the article body, or the first <article>/<main>. */
function articleHtml_(html) {
  const h = html.replace(/<!--[\s\S]*?-->/g, '').replace(/<(nav|footer|aside|form|header)\b[\s\S]*?<\/\1>/gi, '');
  const marks = [/itemprop=["']articleBody["']/i, /class=["'][^"']*\b(entry-content|post-content|article-content|recipe-content|post-body)\b[^"']*["']/i, /<article\b/i, /<main\b/i];
  for (let i = 0; i < marks.length; i++) { const m = h.match(marks[i]); if (m) return h.slice(m.index, m.index + 150000); }
  return h.slice(0, 150000);
}

/** Title without the " - Site Name" tail that sites add to og:title. */
function cleanTitle_(t) { return clean_(t).replace(/\s+[-|–—]\s+[^-|–—]{2,40}$/, '').trim(); }

/** Reads a plain blog post as a recipe. Returns null unless it clearly has an ingredient list and steps. */
function articleRecipe_(html, url, og, site) {
  const title = cleanTitle_(og('og:title') || ((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || ''));
  if (!title) return null;
  let lines = htmlToText_(articleHtml_(html)).split('\n');
  const cut = lines.findIndex(function (l) { return /^(you may also like|related( posts| recipes| articles)?|see more|previous article|leave a (reply|comment)|post navigation|share this)\b/i.test(l.trim()); });
  if (cut >= 0) lines = lines.slice(0, cut);
  const r = parseScan_(title + '\n' + lines.join('\n').replace(/\n{3,}/g, '\n\n'));
  if (r.ingredients.length < 2 || !r.steps.length) return null;
  r.title = title;
  r.summary = og('og:description') || r.summary;
  r.image = og('og:image');
  r.sourceName = site; r.sourceUrl = url;
  r.category = guessCategory_(title + ' ' + r.summary) + extraTags_(title + ' ' + r.summary);
  return r;
}

/** Removes tracking junk (?fbclid=..., utm_*) so the saved source link is clean. */
function stripTracking_(url) {
  const m = String(url).match(/^([^?#]*)(\?[^#]*)?(#.*)?$/);
  if (!m) return url;
  const kept = (m[2] ? m[2].slice(1).split('&') : []).filter(function (p) { return p && !/^(utm_[^=]*|fbclid|gclid|igshid|mc_[^=]*|ref|ref_src)(=|$)/i.test(p); });
  return m[1] + (kept.length ? '?' + kept.join('&') : '');
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
      favorite: String(o.favorite) === 'true', prepTime: o.prepTime, totalTime: o.totalTime, nutrition: o.nutrition, notes: o.notes, cookTime: o.cookTime
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
      lim(r.id, 40), lim(r.title, 200), lim(r.category, 160), lim(r.sourceName, 200), lim(r.sourceUrl, 1000),
      String(Number(r.servings) || 1), lim(r.summary, 2000), lim(r.image, 2000),
      lim(JSON.stringify(r.ingredients || []), 20000), lim(JSON.stringify(r.steps || []), 20000),
      lim(name, 60), new Date().toISOString(),
      String(r.favorite === true), lim(r.prepTime, 40), lim(r.totalTime, 40), lim(r.nutrition, 2000), lim(r.notes, 4000), lim(r.cookTime, 40)
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
