/**
 * Recipe API — paste into Extensions > Apps Script of the recipes Google Sheet.
 * Then: Project Settings > Script properties > add PASSCODE = your shared passcode,
 * and Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The passcode is the only gate, so make it a real one (12+ characters).
 */
const SHEET_NAME = 'Recipes';
const COLS = ['id', 'title', 'category', 'sourceName', 'sourceUrl', 'servings', 'summary',
              'image', 'ingredients', 'steps', 'addedBy', 'added'];

function doGet() { return out_({ ok: true, msg: 'Recipe API is running' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass || String(req.key) !== pass) return out_({ ok: false, error: 'auth' });

  if (req.action === 'list') return out_({ ok: true, recipes: list_() });
  if (req.action === 'save') return save_(req.recipe, req.name);
  return out_({ ok: false, error: 'unknown action' });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
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
      ingredients: parse_(o.ingredients), steps: parse_(o.steps), addedBy: o.addedBy, added: o.added
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
      lim(name, 60), new Date().toISOString()
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
