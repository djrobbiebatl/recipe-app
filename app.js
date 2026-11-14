'use strict';
/* Our Recipes — step 1 shell. Data layer (Store) is swapped for the Google Sheet API in step 2. */

const CATEGORIES = ['Entrée', 'Side', 'Dessert', 'Breakfast', 'Appetizer', 'Drink'];

// ---------- data layer ----------
class AuthError extends Error {}
const Cfg = {
  get url() { return (window.APP_CONFIG && window.APP_CONFIG.apiUrl) || ''; },
  get key() { try { return localStorage.getItem('recipes.key') || ''; } catch (e) { return ''; } },
  get name() { try { return localStorage.getItem('recipes.name') || ''; } catch (e) { return ''; } },
  set(k, n) { try { localStorage.setItem('recipes.key', k); localStorage.setItem('recipes.name', n); } catch (e) {} }
};
const lsGet = k => { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };

async function api(body) {
  if (!Cfg.key) throw new AuthError('no passcode');
  // text/plain keeps this a "simple" request, so the browser skips the CORS preflight Apps Script can't answer.
  const res = await fetch(Cfg.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ key: Cfg.key, ...body }) });
  const data = await res.json();
  if (!data.ok) { if (data.error === 'auth') throw new AuthError('wrong passcode'); throw new Error(data.error || 'server error'); }
  return data;
}

const Store = {
  mem: null,
  async list(fresh) {
    if (!Cfg.url) return lsGet('recipes.local.v1') || SAMPLE;      // demo mode until the Sheet is connected
    if (this.mem && !fresh) return this.mem;
    try {
      this.mem = (await api({ action: 'list' })).recipes;
      lsSet('recipes.cache', this.mem);
    } catch (e) {
      if (e instanceof AuthError) throw e;
      this.mem = lsGet('recipes.cache');
      if (!this.mem) throw e;
      toast('Offline — showing saved recipes');
    }
    return this.mem;
  },
  async get(id) { return (await this.list()).find(r => r.id === id); },
  async remove(id) {
    if (Cfg.url) { await api({ action: 'delete', id }); await this.list(true); return; }
    lsSet('recipes.local.v1', (await this.list()).filter(r => r.id !== id));
  },
  async save(r, quiet) {
    if (Cfg.url) {
      await api({ action: 'save', recipe: r, name: Cfg.name });
      if (quiet) { const i = (this.mem || []).findIndex(x => x.id === r.id); if (i >= 0) { this.mem[i] = r; lsSet('recipes.cache', this.mem); } }
      else await this.list(true);
      return;
    }
    const all = await this.list();
    const i = all.findIndex(x => x.id === r.id);
    if (i >= 0) all[i] = r; else all.unshift(r);
    lsSet('recipes.local.v1', all);
  }
};

const SAMPLE = [
  { id: 's1', title: 'Garlic Butter Chicken Thighs', category: 'Entrée', sourceName: 'Sample Kitchen', sourceUrl: 'https://example.com/chicken',
    servings: 4, summary: 'Crispy-skinned thighs finished in a quick garlic butter pan sauce.', image: '',
    ingredients: [
      { amt: 8, unit: '', item: 'bone-in chicken thighs' }, { amt: 3, unit: 'tbsp', item: 'butter' },
      { amt: 4, unit: 'cloves', item: 'garlic, minced' }, { amt: 0.5, unit: 'cup', item: 'chicken broth' },
      { amt: 1, unit: 'tsp', item: 'smoked paprika' }, { amt: null, unit: '', item: 'salt and pepper' } ],
    steps: ['Season the thighs with salt, pepper and paprika.', 'Sear skin-side down 8 minutes until crisp, flip and cook 5 more.',
      'Remove chicken; add butter and garlic, then broth. Simmer 3 minutes.', 'Return chicken to the pan and spoon sauce over.'] },
  { id: 's2', title: 'Lemon Roasted Broccoli', category: 'Side', sourceName: 'Sample Kitchen', sourceUrl: 'https://example.com/broccoli',
    servings: 4, summary: 'Charred, bright and ready in 20 minutes.', image: '',
    ingredients: [ { amt: 1.5, unit: 'lb', item: 'broccoli crowns' }, { amt: 2, unit: 'tbsp', item: 'olive oil' },
      { amt: 1, unit: '', item: 'lemon, zested and juiced' }, { amt: 0.25, unit: 'cup', item: 'grated parmesan' } ],
    steps: ['Heat oven to 450°F.', 'Toss broccoli with oil and salt; roast 18 minutes.', 'Finish with lemon zest, juice and parmesan.'] },
  { id: 's3', title: 'Brown Butter Chocolate Chip Cookies', category: 'Dessert', sourceName: 'Sample Kitchen', sourceUrl: 'https://example.com/cookies',
    servings: 24, summary: 'Chewy, nutty cookies with crisp edges.', image: '',
    ingredients: [ { amt: 1, unit: 'cup', item: 'butter, browned' }, { amt: 2.25, unit: 'cups', item: 'flour' },
      { amt: 1, unit: 'cup', item: 'brown sugar' }, { amt: 2, unit: '', item: 'eggs' }, { amt: 2, unit: 'cups', item: 'chocolate chips' } ],
    steps: ['Brown the butter and cool 10 minutes.', 'Mix in sugar, eggs, then flour.', 'Fold in chips, scoop, and bake at 350°F for 11 minutes.'] }
];

// ---------- helpers ----------
const $app = document.getElementById('app');
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const safeUrl = u => /^https?:\/\//i.test(u || '') ? u : '';
// Same recipe = same page (ignoring http/https, www, trailing slash, tracking params) or the same title.
const normUrl = u => {
  try {
    const x = new URL(u);
    const keep = [...x.searchParams].filter(([k]) => !/^(utm_|fbclid|gclid|ref$|mc_)/i.test(k)).map(([k, v]) => k + '=' + v).sort().join('&');
    return (x.hostname.replace(/^www\./, '') + x.pathname.replace(/\/+$/, '') + (keep ? '?' + keep : '')).toLowerCase();
  } catch (e) { return ''; }
};
// Returns the matching recipe, tagged by how it matched: 'url' (same page) or 'title' (could be a different recipe).
async function findDup(url, title) {
  const n = normUrl(url), t = (title || '').trim().toLowerCase(), all = await Store.list();
  const byUrl = n && all.find(r => normUrl(r.sourceUrl) === n);
  if (byUrl) return { r: byUrl, by: 'url' };
  const byTitle = t && all.find(r => r.title.trim().toLowerCase() === t);
  return byTitle ? { r: byTitle, by: 'title' } : null;
}
// Photos are shrunk before upload (phone shots are huge); PDFs go up as-is.
const toB64 = blob => new Promise((res, rej) => {
  const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1]); fr.onerror = () => rej(new Error('Could not read the file')); fr.readAsDataURL(blob);
});
async function filePayload(file) {
  if (file.type === 'application/pdf') {
    if (file.size > 8 * 1024 * 1024) throw new Error('That PDF is over 8 MB');
    return { mime: file.type, data: await toB64(file) };
  }
  if (!file.type.startsWith('image/')) throw new Error('Please choose a photo or a PDF');
  let bmp;
  try { bmp = await createImageBitmap(file); } catch (e) { throw new Error('Could not open that image (try a JPEG or PNG)'); }
  const s = Math.min(1, 2000 / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
  return { mime: 'image/jpeg', data: await toB64(blob) };
}
const allCats = list => [...CATEGORIES, ...new Set(list.map(r => r.category).filter(c => c && !CATEGORIES.includes(c)))];
// Star toggle shared by the cards and the recipe page. Updates instantly, saves in the background, rolls back on failure.
async function toggleFav(r, btn) {
  r.favorite = !r.favorite; btn.textContent = r.favorite ? '★' : '☆'; btn.classList.toggle('on', r.favorite);
  try { await Store.save(r, true); }
  catch (e) { r.favorite = !r.favorite; btn.textContent = r.favorite ? '★' : '☆'; btn.classList.toggle('on', r.favorite); toast('Couldn’t update favorite'); }
}
const thumb = u => safeUrl(u) ? ` style="background-image:url('${esc(safeUrl(u))}')"` : '';
const uid = () => 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

function toast(msg) {
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg;
  document.body.appendChild(t); setTimeout(() => t.remove(), 2200);
}

const FRACS = [[1/8, '⅛'], [1/4, '¼'], [1/3, '⅓'], [1/2, '½'], [2/3, '⅔'], [3/4, '¾']];
function fmtAmt(n) {
  if (n == null) return '';
  const whole = Math.floor(n + 1e-6), frac = n - whole;
  if (frac < 0.06) return String(whole || (n < 0.06 ? '0' : ''));
  if (frac > 0.94) return String(whole + 1);
  const f = FRACS.reduce((a, b) => Math.abs(b[0] - frac) < Math.abs(a[0] - frac) ? b : a);
  if (Math.abs(f[0] - frac) < 0.05) return (whole ? whole + ' ' : '') + f[1];
  return String(Math.round(n * 100) / 100);
}

// ---------- views ----------
function settings(msg) {
  $app.innerHTML = `
    <h1>Settings</h1>
    ${msg ? `<div class="note">${esc(msg)}</div>` : ''}
    ${Cfg.url ? '' : '<div class="note">Demo mode: the Google Sheet isn’t connected yet, so recipes only save on this device.</div>'}
    <form id="s">
      <label class="f">Your name (shown as “added by”)</label><input name="name" value="${esc(Cfg.name)}" placeholder="Rob or Daman">
      <label class="f">Shared passcode</label><input name="key" type="password" autocomplete="current-password" value="${esc(Cfg.key)}">
      <p><button class="btn primary" type="submit">Save</button></p>
    </form>`;
  $app.querySelector('#s').onsubmit = e => {
    e.preventDefault(); const d = Object.fromEntries(new FormData(e.target));
    Cfg.set(d.key.trim(), d.name.trim()); Store.mem = null; toast('Saved'); location.hash = '#/';
  };
}

async function home() {
  const all = await Store.list(true);
  let q = '', cat = '';
  $app.innerHTML = `
    ${Cfg.url ? '' : '<div class="note">⚠️ Demo mode — not connected to the Google Sheet. Recipes saved now stay on this device only. Close and reopen the app to load the latest version.</div>'}
    <div class="row"><input id="q" type="search" placeholder="Search recipes or ingredients…" autocomplete="off" style="flex:1;min-width:180px">
      <span class="seg" id="seg"><button class="chip" data-v="tiles">▦ Tiles</button><button class="chip" data-v="list">☰ List</button></span></div>
    <div class="chips" id="chips"></div>
    <div id="grid"></div>`;
  let view = lsGet('recipes.view') === 'list' ? 'list' : 'tiles';
  const star = r => `<button class="star ${r.favorite ? 'on' : ''}" data-id="${esc(r.id)}" title="Favorite" aria-label="Favorite">${r.favorite ? '★' : '☆'}</button>`;
  const draw = () => {
    $app.querySelector('#chips').innerHTML = [['', 'All'], ['__fav', '★ Favorites'], ...allCats(all).map(c => [c, c])].map(([v, label]) =>
      `<button class="chip ${cat === v ? 'on' : ''}" data-c="${esc(v)}">${esc(label)}</button>`).join('');
    $app.querySelectorAll('#seg .chip').forEach(b => b.classList.toggle('on', b.dataset.v === view));
    const hits = all.filter(r => (cat === '__fav' ? r.favorite : (!cat || r.category === cat)) &&
      (!q || (r.title + ' ' + r.summary + ' ' + r.ingredients.map(i => i.item).join(' ')).toLowerCase().includes(q)));
    const g = $app.querySelector('#grid');
    g.className = view === 'list' ? 'list' : 'grid';
    g.innerHTML = !hits.length ? '<p class="muted">No recipes found.</p>' : hits.map(r => view === 'list' ? `
      <a class="lrow" href="#/r/${esc(r.id)}">
        <div class="lthumb"${thumb(r.image)}>${r.image ? '' : '🍽️'}</div>
        <div class="lbody"><b>${esc(r.title)}</b><div class="muted">${esc(r.category)} · serves ${esc(r.servings)}${r.totalTime ? ' · ' + esc(r.totalTime) : ''}</div></div>${star(r)}
      </a>` : `
      <a class="card" href="#/r/${esc(r.id)}">
        <div class="img"${thumb(r.image)}>${r.image ? '' : '🍽️'}${star(r)}</div>
        <div class="body"><h3>${esc(r.title)}</h3><div class="muted">${esc(r.category)} · serves ${esc(r.servings)}</div></div>
      </a>`).join('');
  };
  $app.querySelector('#q').oninput = e => { q = e.target.value.trim().toLowerCase(); draw(); };
  $app.querySelector('#chips').onclick = e => { const b = e.target.closest('[data-c]'); if (b) { cat = b.dataset.c; draw(); } };
  $app.querySelector('#seg').onclick = e => { const b = e.target.closest('[data-v]'); if (b) { view = b.dataset.v; lsSet('recipes.view', view); draw(); } };
  $app.querySelector('#grid').onclick = e => {
    const s = e.target.closest('.star'); if (!s) return;
    e.preventDefault(); e.stopPropagation();
    const r = all.find(x => x.id === s.dataset.id); if (r) toggleFav(r, s).then(() => { if (cat === '__fav') draw(); });
  };
  draw();
}

async function recipe(id) {
  const r = await Store.get(id);
  if (!r) { $app.innerHTML = '<p>Recipe not found. <a href="#/">Back</a></p>'; return; }
  let servings = r.servings || 1;
  const missing = new Set();
  const src = safeUrl(r.sourceUrl);
  $app.innerHTML = `
    <a href="#/" class="noprint">← All recipes</a>
    <h1 class="titlerow"><span>${esc(r.title)}</span><button class="star big noprint ${r.favorite ? 'on' : ''}" id="fav" title="Favorite" aria-label="Favorite">${r.favorite ? '★' : '☆'}</button></h1>
    <div class="muted">${esc(r.category)} · Source: ${src ? `<a href="${esc(src)}" target="_blank" rel="noopener">${esc(r.sourceName || src)}</a>` : esc(r.sourceName)}</div>
    ${r.image ? `<img class="hero" src="${esc(safeUrl(r.image))}" alt="" referrerpolicy="no-referrer">` : ''}
    <p>${esc(r.summary)}</p>
    <div class="row facts">
      ${r.prepTime ? `<span><b>Prep</b> ${esc(r.prepTime)}</span>` : ''}${r.totalTime ? `<span><b>Total</b> ${esc(r.totalTime)}</span>` : ''}
      <span class="servings"><b>Servings</b> <button class="btn noprint" id="minus">−</button><input id="sv" type="number" min="1" value="${servings}"><button class="btn noprint" id="plus">+</button></span>
    </div>
    <div class="row tools">
      <button class="btn" id="print">🖨️ Print</button>
      <button class="btn" id="del">🗑️ Delete</button>
      ${r.image ? '<label class="noprint"><input type="checkbox" id="pimg" checked style="width:auto"> with picture</label>' : ''}
    </div>
    <div class="two">
      <section>
        <h2>Ingredients</h2>
        <p class="muted noprint">Tick what you <b>don’t</b> have.</p>
        <table class="ing" id="ing"></table>
        <div class="row tools"><button class="btn primary" id="shop">Add to shopping list</button></div>
      </section>
      <section><h2>Instructions</h2><ol class="steps">${r.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol></section>
    </div>
    ${r.nutrition ? `<h2>Nutrition</h2><table class="nut">${String(r.nutrition).split('\n').filter(Boolean).map(l => {
      const i = l.indexOf(':'); return i > 0 ? `<tr><td>${esc(l.slice(0, i))}</td><td>${esc(l.slice(i + 1).trim())}</td></tr>` : `<tr><td colspan="2">${esc(l)}</td></tr>`;
    }).join('')}</table>` : ''}`;
  $app.querySelector('#fav').onclick = e => toggleFav(r, e.currentTarget);
  const drawIng = () => {
    const k = servings / (r.servings || 1);
    $app.querySelector('#ing').innerHTML = r.ingredients.map((i, n) =>
      `<tr class="${missing.has(n) ? 'missing' : ''}"><td class="chk"><input type="checkbox" data-n="${n}" ${missing.has(n) ? 'checked' : ''}></td>
       <td class="amt">${esc((fmtAmt(i.amt == null ? null : i.amt * k) + ' ' + (i.unit || '')).trim())}</td><td>${esc(i.item)}</td></tr>`).join('');
  };
  const setServ = v => { servings = Math.max(1, Math.round(v) || 1); $app.querySelector('#sv').value = servings; drawIng(); };
  $app.querySelector('#sv').onchange = e => setServ(+e.target.value);
  $app.querySelector('#minus').onclick = () => setServ(servings - 1);
  $app.querySelector('#plus').onclick = () => setServ(servings + 1);
  $app.querySelector('#ing').onchange = e => {
    const n = +e.target.dataset.n; e.target.checked ? missing.add(n) : missing.delete(n);
    e.target.closest('tr').classList.toggle('missing', e.target.checked);
  };
  $app.querySelector('#print').onclick = () => {
    const p = $app.querySelector('#pimg'); document.body.classList.toggle('noimg', !!p && !p.checked); window.print();
  };
  $app.querySelector('#del').onclick = async () => {
    if (!confirm(`Delete “${r.title}”? This removes it from the shared sheet for everyone and can't be undone.`)) return;
    try { await Store.remove(r.id); toast('Deleted'); location.hash = '#/'; }
    catch (e) { toast('Couldn’t delete: ' + e.message); }
  };
  $app.querySelector('#shop').onclick = async () => {
    const k = servings / (r.servings || 1);
    const items = [...missing].sort((a, b) => a - b).map(n => {
      const i = r.ingredients[n]; return ('• ' + fmtAmt(i.amt == null ? null : i.amt * k) + ' ' + (i.unit || '') + ' ' + i.item).replace(/\s+/g, ' ').replace('• ', '☐ ');
    });
    if (!items.length) return toast('Tick the ingredients you need first');
    const text = items.join('\n');
    // Google Keep has no public API for personal accounts: use the share sheet, then fall back to clipboard.
    try { if (navigator.share) { await navigator.share({ title: 'Shopping list – ' + r.title, text }); return; } } catch (e) { if (e.name === 'AbortError') return; }
    try { await navigator.clipboard.writeText(text); toast('Copied — paste into Google Keep'); } catch (e) { prompt('Copy this list:', text); }
  };
  drawIng();
}

async function add(params) {
  const shared = params || {};
  window.__shared = null; // one-shot, so coming back to Add later doesn't re-import
  $app.innerHTML = `
    <h1>Add a recipe</h1>
    <div class="row"><input id="iu" type="url" placeholder="Paste a recipe link…" value="${esc(shared.url)}" style="flex:1;min-width:200px">
      <button class="btn primary" id="imp" type="button">Import</button></div>
    <div class="row">
      <button class="btn" id="scan" type="button">📷 Scan photo or PDF</button>
      <input id="file" type="file" accept="image/*,application/pdf" hidden>
      <span class="muted">…or drag a photo / PDF anywhere on this page</span>
    </div>
    <p id="impmsg" class="muted"></p>
    <details id="raw" class="note" hidden><summary>Show the scanned text</summary><pre id="rawtxt" style="white-space:pre-wrap"></pre></details>
    <form id="f">
      <label class="f">Title</label><input name="title" required value="${esc(shared.title)}">
      <label class="f">Category</label>
      <div class="row" style="margin-top:0"><select name="category" style="flex:1">${allCats(await Store.list()).map(c => `<option>${esc(c)}</option>`).join('')}</select>
        <button class="btn" id="newcat" type="button">+ New category</button></div>
      <div class="cols3">
        <div><label class="f">Servings</label><input name="servings" type="number" min="1" value="4"></div>
        <div><label class="f">Prep time</label><input name="prepTime" placeholder="15 min"></div>
        <div><label class="f">Total time</label><input name="totalTime" placeholder="1 hr"></div>
      </div>
      <label class="f">Picture</label>
      <div class="row" style="margin-top:0"><input name="image" type="url" placeholder="Picture link, or upload one →" style="flex:1;min-width:180px">
        <button class="btn" id="pic" type="button">⬆️ Upload photo</button><input id="picfile" type="file" accept="image/*" hidden></div>
      <img id="prev" class="hero" alt="" hidden style="max-height:180px">
      <label class="f">Summary</label><textarea name="summary" rows="2"></textarea>
      <label class="f">Ingredients — one per line, like “2 cups flour”</label><textarea name="ingredients" rows="6" required></textarea>
      <label class="f">Steps — one per line</label><textarea name="steps" rows="6" required></textarea>
      <label class="f">Nutrition (optional) — one per line, like “Calories: 250”</label><textarea name="nutrition" rows="4"></textarea>
      <label class="f">Source name</label><input name="sourceName">
      <label class="f">Source link</label><input name="sourceUrl" type="url" value="${esc(shared.url)}">
      <p><button class="btn primary" type="submit">Save recipe</button></p>
    </form>`;
  const f = $app.querySelector('#f'), msg = $app.querySelector('#impmsg'), impBtn = $app.querySelector('#imp');
  const runImport = async () => {
    const url = $app.querySelector('#iu').value.trim();
    if (!url) return;
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — import needs it.'; return; }
    msg.textContent = 'Reading the recipe…'; impBtn.disabled = true;
    const showDup = d => { msg.innerHTML = `You already have this one: <a href="#/r/${esc(d.id)}">${esc(d.title)}</a>`; };
    try {
      let dup = await findDup(url, '');
      if (dup) return showDup(dup.r);                       // same page: don't import twice
      const r = (await api({ action: 'import', url })).recipe;
      dup = await findDup(r.sourceUrl || url, r.title);
      if (dup && dup.by === 'url') return showDup(dup.r);
      if (dup && !confirm(`You already have a recipe called “${dup.r.title}”. Is this a different one? OK = import it anyway, Cancel = skip.`)) return showDup(dup.r);
      const found = r.ingredients && r.ingredients.length;
      const set = (k, v) => { if (v != null && v !== '') f.elements[k].value = v; };
      set('title', r.title); set('sourceName', r.sourceName); set('sourceUrl', r.sourceUrl || url);
      set('servings', r.servings); set('image', r.image); set('summary', r.summary);
      set('prepTime', r.prepTime); set('totalTime', r.totalTime); set('nutrition', r.nutrition);
      if (r.image) showPrev(r.image);
      if ([...f.elements.category.options].some(o => o.value === r.category)) set('category', r.category);
      set('ingredients', (r.ingredients || []).join('\n')); set('steps', (r.steps || []).join('\n'));
      msg.textContent = found ? 'Imported — check it over, then save.'
        : 'That page has no recipe data I can read. I filled in what I could — add the rest by hand.';
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t import: ' + e.message;
    } finally { impBtn.disabled = false; }
  };
  impBtn.onclick = runImport;
  if (shared.url) runImport();

  // ----- new category, recipe picture -----
  $app.querySelector('#newcat').onclick = () => {
    const n = (prompt('New category name:') || '').trim().slice(0, 30);
    if (!n) return;
    const sel = f.elements.category;
    let o = [...sel.options].find(x => x.value.toLowerCase() === n.toLowerCase());
    if (!o) { o = new Option(n, n); sel.add(o); }
    sel.value = o.value;
  };
  const prev = $app.querySelector('#prev');
  const showPrev = u => { if (safeUrl(u)) { prev.referrerPolicy = 'no-referrer'; prev.src = safeUrl(u); prev.hidden = false; } else prev.hidden = true; };
  f.elements.image.oninput = () => showPrev(f.elements.image.value.trim());
  const picFile = $app.querySelector('#picfile');
  $app.querySelector('#pic').onclick = () => picFile.click();
  picFile.onchange = async () => {
    const file = picFile.files[0]; picFile.value = '';
    if (!file) return;
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — photo upload needs it.'; return; }
    msg.textContent = 'Uploading photo…';
    try {
      const p = await filePayload(file);
      if (p.mime === 'application/pdf') throw new Error('Please choose a photo, not a PDF');
      const d = await api({ action: 'upload', mime: p.mime, data: p.data });
      f.elements.image.value = d.imageUrl; showPrev(d.imageUrl); msg.textContent = 'Photo added.';
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t upload: ' + e.message;
    }
  };

  // ----- photo / PDF scan (Google Drive OCR, done in the Apps Script) -----
  const fileEl = $app.querySelector('#file');
  const runScan = async file => {
    if (!file) return;
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — scanning needs it.'; return; }
    msg.textContent = 'Scanning… this can take 10–30 seconds.';
    try {
      const p = await filePayload(file);
      const d = await api({ action: 'scan', mime: p.mime, data: p.data });
      const r = d.recipe;
      const dup = await findDup('', r.title);
      if (dup && !confirm(`You already have a recipe called “${dup.r.title}”. Is this a different one? OK = fill the form anyway, Cancel = skip.`)) {
        msg.innerHTML = `You already have: <a href="#/r/${esc(dup.r.id)}">${esc(dup.r.title)}</a>`; return;
      }
      const set = (k, v) => { if (v != null && v !== '') f.elements[k].value = v; };
      set('title', r.title); set('servings', r.servings); set('summary', r.summary);
      set('prepTime', r.prepTime); set('totalTime', r.totalTime); set('nutrition', r.nutrition);
      if ([...f.elements.category.options].some(o => o.value === r.category)) set('category', r.category);
      set('ingredients', (r.ingredients || []).join('\n')); set('steps', (r.steps || []).join('\n'));
      // An uploaded/scanned source has no web page: link the saved original and mark it as an import.
      if (d.fileUrl) { set('sourceUrl', d.fileUrl); set('sourceName', 'Import'); }
      $app.querySelector('#rawtxt').textContent = d.text || ''; $app.querySelector('#raw').hidden = !d.text;
      msg.textContent = 'Scanned — scans are rarely perfect, so check every line before saving.';
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t scan: ' + e.message;
    }
  };
  $app.querySelector('#scan').onclick = () => fileEl.click();
  fileEl.onchange = () => { runScan(fileEl.files[0]); fileEl.value = ''; };
  const stop = e => { e.preventDefault(); };
  document.ondragover = stop;
  document.ondrop = e => { e.preventDefault(); if (location.hash.startsWith('#/add')) runScan(e.dataTransfer.files[0]); };
  f.onsubmit = async e => {
    e.preventDefault();
    const d = Object.fromEntries(new FormData(e.target));
    const rec = { id: uid(), title: d.title.trim(), category: d.category, sourceName: d.sourceName.trim(), sourceUrl: d.sourceUrl.trim(),
      servings: +d.servings || 1, summary: d.summary.trim(), image: d.image.trim(), favorite: false,
      prepTime: d.prepTime.trim(), totalTime: d.totalTime.trim(), nutrition: d.nutrition.trim(),
      ingredients: d.ingredients.split('\n').map(parseIngredient).filter(Boolean),
      steps: d.steps.split('\n').map(s => s.replace(/^\s*\d{1,2}[.)]\s+/, '').trim()).filter(Boolean) };
    const btn = e.target.querySelector('button[type=submit]'); btn.disabled = true; btn.textContent = 'Saving…';
    try {
      const dup = await findDup(rec.sourceUrl, rec.title);
      if (dup && !confirm(`“${dup.r.title}” is already in your recipes. Save another copy anyway?`)) { btn.disabled = false; btn.textContent = 'Save recipe'; return; }
      await Store.save(rec); location.hash = '#/r/' + rec.id; }
    catch (err) {
      btn.disabled = false; btn.textContent = 'Save recipe';
      if (err instanceof AuthError) return settings('That passcode didn’t work. Check it and try again.');
      toast('Couldn’t save: ' + err.message);
    }
  };
}

const UNICODE_FRACS = { '½': .5, '⅓': 1/3, '⅔': 2/3, '¼': .25, '¾': .75, '⅛': .125 };
function parseIngredient(line) {
  line = line.trim(); if (!line) return null;
  const m = line.match(/^(\d+\s+\d+\/\d+|\d+\/\d+|\d*\.?\d+\s*[½⅓⅔¼¾⅛]?|[½⅓⅔¼¾⅛])\s*(.*)$/);
  if (!m) return { amt: null, unit: '', item: line };
  const raw = m[1].trim(); let amt = 0;
  for (const part of raw.replace(/([½⅓⅔¼¾⅛])/g, ' $1').trim().split(/\s+/)) {
    amt += UNICODE_FRACS[part] ?? (part.includes('/') ? part.split('/').reduce((a, b) => a / b) : parseFloat(part));
  }
  const rest = m[2].match(/^(cups?|tbsp|tsp|tablespoons?|teaspoons?|oz|ounces?|lbs?|pounds?|g|grams?|kg|ml|l|cloves?|cans?|pinch|dash|slices?)\b\.?\s*(.*)$/i);
  return rest ? { amt, unit: rest[1], item: rest[2] } : { amt, unit: '', item: m[2] };
}

// ---------- router ----------
async function route() {
  const h = location.hash || '#/';
  window.scrollTo(0, 0);
  try {
    if (h.startsWith('#/settings')) return settings();
    if (h.startsWith('#/r/')) return await recipe(decodeURIComponent(h.slice(4)));
    if (h.startsWith('#/add')) return await add(window.__shared);
    return await home();
  } catch (e) {
    if (e instanceof AuthError) return settings(Cfg.key ? 'That passcode didn’t work.' : 'Enter the shared passcode to open the recipes.');
    $app.innerHTML = `<p>Couldn’t load recipes: ${esc(e.message)}. <a href="#/">Try again</a></p>`;
  }
}
window.addEventListener('hashchange', route);

// Share-target landing: ?title=&text=&url= → open new recipe entry
(function () {
  const p = new URLSearchParams(location.search);
  if (p.has('url') || p.has('text') || p.has('title')) {
    let url = p.get('url') || '';
    if (!url) { const m = (p.get('text') || '').match(/https?:\/\/\S+/); url = m ? m[0] : ''; }
    window.__shared = { title: p.get('title') || '', url };
    history.replaceState(null, '', location.pathname + '#/add');
  }
})();

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
route();
