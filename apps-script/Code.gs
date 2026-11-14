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
  if (req.action === 'import') return import_(req.url);
  return out_({ ok: false, error: 'unknown action' });
}

// ---------- link import (runs here because browsers can't read other sites' pages) ----------
function import_(url) {
  if (!/^https?:\/\/[^\s]+$/i.test(String(url || ''))) return out_({ ok: false, error: 'That doesn\'t look like a web link' });
  let html;
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; RecipeApp/1.0)', 'Accept-Language': 'en-US,en;q=0.9' }
    });
    if (res.getResponseCode() >= 400) return out_({ ok: false, error: 'The site returned error ' + res.getResponseCode() });
    html = res.getContentText();
  } catch (e) { return out_({ ok: false, error: 'Could not reach that page' }); }
  const r = extract_(html, url);
  return out_({ ok: true, found: r.found, recipe: r.recipe });
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
    ingredients: [].concat(node.recipeIngredient || node.ingredients || []).map(clean_).filter(String),
    steps: steps_(node.recipeInstructions)
  } };
}

function findRecipe_(x) {
  if (!x || typeof x !== 'object') return null;
  if (Array.isArray(x)) { for (let i = 0; i < x.length; i++) { const r = findRecipe_(x[i]); if (r) return r; } return null; }
  const t = [].concat(x['@type'] || []);
  if (t.indexOf('Recipe') >= 0) return x;
  return findRecipe_(x['@graph']) || findRecipe_(x.mainEntity) || findRecipe_(x.mainEntityOfPage);
}

function steps_(ins) {
  if (!ins) return [];
  if (typeof ins === 'string') return clean_(ins.replace(/<\/?(li|p|br)[^>]*>/gi, '\n')).split(/\n+/).map(function (s) { return s.trim(); }).filter(String);
  const out = [];
  [].concat(ins).forEach(function (s) {
    if (typeof s === 'string') { const t = clean_(s); if (t) out.push(t); }
    else if (s && s.itemListElement) steps_(s.itemListElement).forEach(function (t) { out.push(t); });
    else if (s && (s.text || s.name)) { const t = clean_(s.text || s.name); if (t) out.push(t); }
  });
  return out;
}

function guessCategory_(s) {
  s = String(s).toLowerCase();
  if (/dessert|cookie|cake|pie\b|brownie|ice cream|cobbler/.test(s)) return 'Dessert';
  if (/breakfast|brunch|pancake|waffle|omelet|oatmeal/.test(s)) return 'Breakfast';
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
