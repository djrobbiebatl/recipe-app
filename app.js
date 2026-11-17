'use strict';
/* Our Recipes — step 1 shell. Data layer (Store) is swapped for the Google Sheet API in step 2. */

const CATEGORIES = ['Entrée', 'Side', 'Dessert', 'Breakfast', 'Appetizer', 'Drink', 'Crock Pot', 'Instant Pot'];
// A recipe can carry several categories/tags. They live in the one "category" cell as "Side, Breakfast".
const tagsOf = r => String(r.category || '').split(',').map(s => s.trim()).filter(Boolean);

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
  const quick = ['list', 'save', 'delete', 'websuggest', 'planlist', 'planadd', 'plandelete', 'shoplist', 'shopadd', 'shopset', 'shopdelete'].includes(body.action);
  const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), quick ? 40000 : 120000);
  let data;
  try {
    const res = await fetch(Cfg.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, signal: ctl.signal,
      body: JSON.stringify({ key: Cfg.key, ...body }) });
    data = await res.json();
  } catch (e) {
    throw new Error(e.name === 'AbortError' ? 'The server took too long to answer' : 'Could not reach the server (' + e.message + ')');
  } finally { clearTimeout(timer); }
  if (!data.ok) { if (data.error === 'auth') throw new AuthError('wrong passcode'); const err = new Error(data.error || 'server error'); err.code = data.code; err.quota = data.quota; throw err; }
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
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  const DOC = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', doc: 'application/msword', txt: 'text/plain', eml: 'message/rfc822' };
  if (DOC[ext] || file.type === 'text/plain') {
    if (file.size > 8 * 1024 * 1024) throw new Error('That file is over 8 MB');
    return { mime: DOC[ext] || file.type, data: await toB64(file) };
  }
  if (!file.type.startsWith('image/')) throw new Error('Please choose a photo, PDF, Word document or text file');
  let bmp;
  try { bmp = await createImageBitmap(file); } catch (e) { throw new Error('Could not open that image (try a JPEG or PNG)'); }
  const s = Math.min(1, 2000 / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
  return { mime: 'image/jpeg', data: await toB64(blob) };
}
const allCats = list => [...CATEGORIES, ...new Set(list.flatMap(tagsOf).filter(c => !CATEGORIES.includes(c)))];
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
  // Show saved recipes instantly, then refresh from the sheet in the background.
  let all = Cfg.url ? (Store.mem || lsGet('recipes.cache') || null) : await Store.list();
  let loading = !all; if (!all) all = [];
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
    const hits = all.filter(r => (cat === '__fav' ? r.favorite : (!cat || tagsOf(r).includes(cat))) &&
      (!q || (r.title + ' ' + r.summary + ' ' + r.ingredients.map(i => i.item).join(' ')).toLowerCase().includes(q)));
    const g = $app.querySelector('#grid');
    g.className = view === 'list' ? 'list' : 'grid';
    g.innerHTML = !hits.length ? `<p class="muted">${loading ? 'Loading recipes… the first load can take up to 20 seconds.' : 'No recipes found.'}</p>` : hits.map(r => view === 'list' ? `
      <a class="lrow" href="#/r/${esc(r.id)}">
        <div class="lthumb"${thumb(r.image)}>${r.image ? '' : '🍽️'}</div>
        <div class="lbody"><b>${esc(r.title)}</b><div class="muted">${esc(tagsOf(r).join(' · '))} · serves ${esc(r.servings)}</div>${timesLine(r) ? `<div class="muted">${esc(timesLine(r))}</div>` : ''}</div>${star(r)}
      </a>` : `
      <a class="card" href="#/r/${esc(r.id)}">
        <div class="img"${thumb(r.image)}>${r.image ? '' : '🍽️'}${star(r)}</div>
        <div class="body"><h3>${esc(r.title)}</h3><div class="muted">${esc(tagsOf(r).join(' · '))} · serves ${esc(r.servings)}</div></div>
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
  if (Cfg.url) {
    Store.list(true).then(fresh => { all = fresh; loading = false; if ($app.querySelector('#grid')) draw(); })
      .catch(e => {
        loading = false;
        if (e instanceof AuthError) return settings(Cfg.key ? 'That passcode didn’t work.' : 'Enter the shared passcode to open the recipes.');
        const g = $app.querySelector('#grid');
        if (g) g.innerHTML = `<p>Couldn’t load recipes: ${esc(e.message)}. <a href="#/">Try again</a></p>`;
      });
  }
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
    <div class="muted">${esc(tagsOf(r).join(' · '))}<span class="noprint"> · Source: ${src ? `<a href="${esc(src)}" target="_blank" rel="noopener">${esc(r.sourceName || src)}</a>` : esc(r.sourceName)}</span></div>
    ${r.image ? `<img class="hero" src="${esc(safeUrl(r.image))}" alt="" referrerpolicy="no-referrer">` : ''}
    <p>${esc(r.summary)}</p>
    <div class="row facts">
      ${r.prepTime ? `<span><b>Prep</b> ${esc(r.prepTime)}</span>` : ''}${r.cookTime ? `<span><b>Cook</b> ${esc(r.cookTime)}</span>` : ''}${r.totalTime ? `<span><b>Total</b> ${esc(r.totalTime)}</span>` : ''}
      <span class="servings"><b>Servings</b> <button class="btn noprint" id="minus">−</button><input id="sv" type="number" min="1" value="${servings}"><button class="btn noprint" id="plus">+</button></span>
    </div>
    <div class="row tools">
      <button class="btn" id="print">🖨️ Print</button>
      ${r.image ? '<label class="noprint"><input type="checkbox" id="pimg" style="width:auto"> print with picture</label>' : ''}
    </div>
    <div class="row tools">
      <a class="btn" href="#/edit/${esc(r.id)}">✏️ Edit</a>
      <button class="btn" id="del">🗑️ Delete</button>
    </div>
    <div class="row tools">
      <a class="btn" target="_blank" rel="noopener" href="https://www.google.com/search?q=${encodeURIComponent('recipes similar to ' + r.title)}">🔍 More like this</a>
      <span class="noprint cal"><button class="btn" id="plan">📅 Add to calendar</button></span>
    </div>
    ${preheatLine(r.steps) ? `<p class="preheat">🔥 ${esc(preheatLine(r.steps))}</p>` : ''}
    ${marinateLine(r.steps) ? `<p class="preheat">⏱️ ${esc(marinateLine(r.steps))}</p>` : ''}
    <div class="two">
      <section>
        <h2>Ingredients</h2>
        <p class="muted noprint">Tick what you <b>don’t</b> have.</p>
        <table class="ing" id="ing"></table>
        <div class="row tools"><button class="btn primary" id="weekly">🛒 Add to weekly list</button><button class="btn" id="shop">Send to Keep / copy</button></div>
      </section>
      <section><h2>Instructions <label class="noprint" id="awakeRow" hidden style="font-size:.9rem;font-weight:400"><input type="checkbox" id="awake" style="width:auto"> keep screen on</label></h2><ol class="steps">${r.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol></section>
    </div>
    ${r.notes ? `<h2>Notes</h2><p class="notes">${esc(r.notes)}</p>` : ''}
    ${r.nutrition ? `<h2>Nutrition</h2><table class="nut">${String(r.nutrition).split('\n').filter(Boolean).map(l => {
      const i = l.indexOf(':'); return i > 0 ? `<tr><td>${esc(l.slice(0, i))}</td><td>${esc(l.slice(i + 1).trim())}</td></tr>` : `<tr><td colspan="2">${esc(l)}</td></tr>`;
    }).join('')}</table>` : ''}`;
  $app.querySelector('#fav').onclick = e => toggleFav(r, e.currentTarget);
  const drawIng = () => {
    const k = servings / (r.servings || 1);
    $app.querySelector('#ing').innerHTML = r.ingredients.map((i, n) =>
      `<tr class="${missing.has(n) ? 'missing' : ''}"><td class="chk"><input type="checkbox" data-n="${n}" ${missing.has(n) ? 'checked' : ''}></td>
       <td class="amt">${esc(amtText(i, k))}</td><td>${esc(i.item)}</td></tr>`).join('');
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
    const p = $app.querySelector('#pimg'); document.body.classList.toggle('noimg', !!p && !p.checked);
    window.addEventListener('afterprint', () => document.body.classList.remove('noimg'), { once: true });
    window.print();
  };
  $app.querySelector('#del').onclick = async () => {
    if (!confirm(`Delete “${r.title}”? This removes it from the shared sheet for everyone and can't be undone.`)) return;
    try { await Store.remove(r.id); toast('Deleted'); location.hash = '#/'; }
    catch (e) { toast('Couldn’t delete: ' + e.message); }
  };
  // Keep the screen on while cooking (Screen Wake Lock). Phones drop the lock when the app is hidden, so ask again on return.
  if ('wakeLock' in navigator) {
    const box = $app.querySelector('#awake'); $app.querySelector('#awakeRow').hidden = false;
    let lock = null;
    const grab = async () => {
      try { lock = await navigator.wakeLock.request('screen'); lock.addEventListener('release', () => { lock = null; }); }
      catch (e) { box.checked = false; toast('Couldn’t keep the screen on'); }
    };
    const drop = () => { try { if (lock) lock.release(); } catch (e) {} lock = null; };
    box.onchange = () => box.checked ? grab() : drop();
    const back = () => { if (!document.body.contains(box)) { document.removeEventListener('visibilitychange', back); drop(); return; } if (document.visibilityState === 'visible' && box.checked && !lock) grab(); };
    document.addEventListener('visibilitychange', back);
    // leaving this recipe releases the lock
    window.addEventListener('hashchange', () => { drop(); document.removeEventListener('visibilitychange', back); }, { once: true });
  }
  // One button: tapping it opens the date picker; choosing a date adds the recipe to that night.
  // (Our own day picker: phones' built-in date inputs can fire immediately with today's date.)
  $app.querySelector('#plan').onclick = () => pickDay(async date => {
    const b = $app.querySelector('#plan'); b.disabled = true;
    try {
      const d = await api({ action: 'planadd', date, recipeId: r.id, title: r.title });
      toast(d.already ? 'Already on the calendar for that night' : 'Added to ' + niceDate(date));
    } catch (err) { toast('Couldn’t add: ' + err.message); }
    b.disabled = false;
  });
  const chosen = () => {
    const k = servings / (r.servings || 1);
    return [...missing].sort((a, b) => a - b).map(n => { const i = r.ingredients[n]; return { amt: i.amt == null ? null : Math.round(i.amt * k * 1000) / 1000, unit: i.unit || '', item: i.item }; });
  };
  $app.querySelector('#weekly').onclick = async e => {
    const items = chosen();
    if (!items.length) return toast('Tick the ingredients you need first');
    const b = e.currentTarget; b.disabled = true;
    try {
      const d = await api({ action: 'shopadd', items, recipe: r.title });
      toast(`Added to the weekly list${d.merged ? ' (' + d.merged + ' combined with items already there)' : ''}`);
      missing.clear(); drawIng();
    } catch (err) { toast('Couldn’t add: ' + err.message); }
    b.disabled = false;
  };
  $app.querySelector('#shop').onclick = async () => {
    const k = servings / (r.servings || 1);
    const items = [...missing].sort((a, b) => a - b).map(n => {
      const i = r.ingredients[n]; return ('• ' + fmtAmt(i.amt == null ? null : i.amt * k) + ' ' + (i.unit || '') + ' ' + i.item).replace(/\s+/g, ' ').replace('• ', '☐ ');
    });
    if (!items.length) return toast('Tick the ingredients you need first');
    const text = items.join('\n');
    // Phones/tablets: the share sheet (pick Google Keep). Computers: copy, because Windows' share box only offers
    // Nearby Sharing / Outlook / Copilot. Google Keep has no public API for personal accounts.
    const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Macintosh/i.test(navigator.userAgent));
    if (isMobile && navigator.share) {
      try { await navigator.share({ title: 'Shopping list – ' + r.title, text }); return; } catch (e) { if (e.name === 'AbortError') return; }
    }
    try { await navigator.clipboard.writeText(text); toast('List copied — paste it into an email or note'); } catch (e) { prompt('Copy this list:', text); }
  };
  drawIng();
}

// Pulls "Preheat the oven to 425°F." style sentences out of the steps so they can sit above the ingredients.
function preheatLine(steps) {
  const found = [];
  for (const s of steps || []) {
    for (const sentence of String(s).split(/(?<=[.!?])\s+/)) {
      if (/\bpre-?heat(ed|ing)?\b|\bheat (the |your )?(oven|grill|air fryer|broiler)\b/i.test(sentence) && !found.includes(sentence.trim())) found.push(sentence.trim());
    }
  }
  return found.slice(0, 2).join(' ');
}
// "Marinate for at least 2 hours." — only sentences that actually give a time, so "whisk the marinade" doesn't count.
function marinateLine(steps) {
  const found = [];
  for (const s of steps || []) {
    for (const sentence of String(s).split(/(?<=[.!?])\s+/)) {
      if (/\bmarinat(e|es|ed|ing)\b/i.test(sentence) && /\b(hours?|hrs?|minutes?|mins?|overnight|days?)\b/i.test(sentence) && !found.includes(sentence.trim())) found.push(sentence.trim());
    }
  }
  return found.slice(0, 2).join(' ');
}

const timesLine = r => [r.prepTime && 'Prep ' + r.prepTime, r.cookTime && 'Cook ' + r.cookTime, r.totalTime && 'Total ' + r.totalTime].filter(Boolean).join(' · ');
// The Amount column. No quantity at all means "season it yourself": seasonings read "To taste", anything else "Desired".
const SEASONING = /\b(salt|pepper|seasoning|spices?|powder|paprika|cumin|oregano|basil|thyme|parsley|rosemary|sage|cinnamon|nutmeg|garlic|herbs?|cayenne|chili flakes|red pepper flakes|bay leaf|vanilla|worcestershire|hot sauce)\b/i;
function amtText(i, k) {
  const a = ((i.amt == null ? '' : fmtAmt(i.amt * k)) + ' ' + (i.unit || '')).trim();
  if (a) return a;
  return SEASONING.test(i.item) || /to taste/i.test(i.item) ? 'To taste' : 'Desired';
}
const ingText = i => ((i.amt == null ? '' : fmtAmt(i.amt)) + ' ' + (i.unit || '') + ' ' + i.item).replace(/\s+/g, ' ').trim();

// One screen for both adding (ex = undefined) and editing (ex = the saved recipe).
// ---------- Review tab: recipes forwarded by email wait here until you tidy them up and post them ----------
function setBadge(n) {
  const b = document.getElementById('badge');
  try { localStorage.setItem('recipes.badge', String(n)); } catch (e) {}
  b.textContent = n; b.hidden = !n;
}
async function refreshBadge() {
  if (!Cfg.url || !Cfg.key) return;
  try { setBadge((await api({ action: 'intakecount' })).count); } catch (e) { /* the badge is a nicety */ }
}

async function review() {
  $app.innerHTML = `<h1>Review</h1>
    <p class="muted">Forward a recipe email to the app’s address and it waits here. Open it, clean it up, then post it to your recipes.</p>
    <div id="rv"><p class="muted">Checking for new emails… this can take a few seconds.</p></div>`;
  const rv = $app.querySelector('#rv');
  if (!Cfg.url) { rv.textContent = 'Connect the Google Sheet first — email review needs it.'; return; }
  let d;
  try { d = await api({ action: 'intakesync' }); }
  catch (e) {
    if (e instanceof AuthError) throw e;
    rv.innerHTML = `<p>${esc(e.message)}</p>`; return;
  }
  setBadge(d.items.length);
  const draw = () => {
    rv.innerHTML = (d.address ? `<div class="note">Forward recipes to: <b id="addr">${esc(d.address)}</b> <button class="btn" id="copyaddr" type="button">Copy</button>
        <div class="muted">Anyone who has this address can send recipes here. Nothing goes live until you post it.</div></div>`
        : '<div class="note">The email address will show here after the intake account sends its first email to the app. See the setup steps at the top of <b>Intake.gs</b> in the project.</div>') +
      (d.items.length ? d.items.map(i => `
        <div class="qitem" data-id="${esc(i.id)}">
          <b>${esc(i.subject || '(no subject)')}</b>
          <div class="muted">From ${esc(i.from || 'unknown')} · ${esc(new Date(i.date).toLocaleDateString())}</div>
          <div class="muted">${esc(i.snippet)}</div>
          <div class="row"><button class="btn primary" data-act="open" type="button">Review &amp; post</button>
            <button class="btn" data-act="dismiss" type="button">Dismiss</button></div>
        </div>`).join('') : '<p>Nothing waiting. 🎉 Forward a recipe email and it will show up here.</p>');
    const c = rv.querySelector('#copyaddr');
    if (c) c.onclick = async () => { try { await navigator.clipboard.writeText(d.address); toast('Address copied'); } catch (e) { prompt('Copy this address:', d.address); } };
  };
  draw();
  rv.onclick = async e => {
    const btn = e.target.closest('[data-act]'); if (!btn) return;
    const id = btn.closest('.qitem').dataset.id;
    btn.disabled = true;
    try {
      if (btn.dataset.act === 'dismiss') {
        if (!confirm('Dismiss this email? It leaves the queue and won’t come back.')) { btn.disabled = false; return; }
        await api({ action: 'intakemark', id, status: 'dismissed' });
        d.items = d.items.filter(x => x.id !== id); setBadge(d.items.length); draw();
      } else {
        const data = await api({ action: 'intakeget', id });
        window.__intake = { id, data }; location.hash = '#/add';
      }
    } catch (err) {
      btn.disabled = false;
      if (err instanceof AuthError) return settings('That passcode didn’t work.');
      toast('Couldn’t do that: ' + err.message);
    }
  };
}

// Google search results / AI overview share links are pages built by script; they hold no recipe we can read.
function isGooglePage(u) {
  try {
    const x = new URL(u), h = x.hostname.toLowerCase();
    return h === 'share.google' || h === 'g.co' || (/(^|\.)google\.[a-z.]+$/.test(h) && /^\/(search|url|aclk|share)/.test(x.pathname));
  } catch (e) { return false; }
}

async function add(params, ex) {
  const shared = params || {};
  window.__shared = null; // one-shot, so coming back to Add later doesn't re-import
  const intake = window.__intake || null; window.__intake = null;   // set when opened from the Review tab
  const cancelHref = ex ? '#/r/' + esc(ex.id) : intake ? '#/review' : '#/';
  const v = ex || {};
  const cats = allCats(await Store.list());
  tagsOf(v).forEach(t => { if (!cats.includes(t)) cats.push(t); });
  const picked = new Set(ex ? tagsOf(ex) : ['Entrée']);   // a new recipe starts as Entrée; tap tags to add or remove
  $app.innerHTML = `
    <h1>${ex ? 'Edit recipe' : 'Add a recipe'}</h1>
    <div id="importbox"${ex ? ' hidden' : ''}>
    <div class="row"><input id="iu" type="url" placeholder="Paste a recipe link…" value="${esc(shared.url)}" style="flex:1;min-width:200px">
      <button class="btn primary" id="imp" type="button">Import</button></div>
    <div class="row">
      <button class="btn" id="scan" type="button">📷 Scan photo or file</button>
      <button class="btn" id="drv" type="button">☁️ Import from Drive</button>
      <input id="file" type="file" accept="image/*,application/pdf,.pdf,.doc,.docx,.txt,.eml,text/plain" hidden>
      <span class="muted">Photo, PDF, Word, .txt or email (.eml) — or drag a file onto this page. You can also paste a Drive link above.</span>
    </div>
    <div id="drvbox" class="note" hidden></div>
    </div>
    <p id="impmsg" class="muted"></p>
    <details id="raw" class="note" hidden><summary>Show the scanned text</summary><pre id="rawtxt" style="white-space:pre-wrap"></pre></details>
    <form id="f">
      <p class="savetop"><button class="btn primary" type="submit">${ex ? 'Save changes' : 'Save recipe'}</button>
        <a class="btn cancel" href="${cancelHref}">Cancel</a></p>
      <label class="f">Title</label><input name="title" required value="${esc(ex ? v.title : shared.title)}">
      <label class="f">Categories</label>
      <div class="muted" style="margin:-2px 0 6px">Tap every one that applies — a recipe can have several.</div>
      <div class="chips" id="tagchips" style="margin:0 0 4px"></div>
      <input type="hidden" name="category" value="${esc([...picked].join(', '))}">
      <button class="btn" id="newcat" type="button">+ New category</button>
      <div class="cols4">
        <div><label class="f">Servings</label><input name="servings" type="number" min="1" value="${esc(v.servings || 4)}"></div>
        <div><label class="f">Prep time</label><input name="prepTime" placeholder="15 min" value="${esc(v.prepTime)}"></div>
        <div><label class="f">Cook time</label><input name="cookTime" placeholder="30 min" value="${esc(v.cookTime)}"></div>
        <div><label class="f">Total time</label><input name="totalTime" placeholder="1 hr" value="${esc(v.totalTime)}"></div>
      </div>
      <label class="f">Picture</label>
      <div class="row" style="margin-top:0"><input name="image" type="url" placeholder="Picture link, or upload one →" style="flex:1;min-width:180px" value="${esc(v.image)}">
        <button class="btn" id="pic" type="button">⬆️ Upload photo</button><input id="picfile" type="file" accept="image/*" hidden></div>
      <img id="prev" class="hero" alt="" hidden style="max-height:180px">
      <label class="f">Summary</label><textarea name="summary" rows="2">${esc(v.summary)}</textarea>
      <label class="f">Ingredients — one per line, like “2 cups flour”</label><textarea name="ingredients" rows="6" required>${esc((v.ingredients || []).map(ingText).join('\n'))}</textarea>
      <label class="f">Steps — one per line</label><textarea name="steps" rows="6" required>${esc((v.steps || []).join('\n'))}</textarea>
      <label class="f">Notes (optional) — tips, changes you made, who loved it</label><textarea name="notes" rows="4">${esc(v.notes)}</textarea>
      <label class="f">Nutrition (optional) — one per line, like “Calories: 250”</label><textarea name="nutrition" rows="4">${esc(v.nutrition)}</textarea>
      <label class="f">Source name</label><input name="sourceName" value="${esc(v.sourceName)}">
      <label class="f">Source link</label>
      <div class="row" style="margin-top:0"><input name="sourceUrl" type="url" value="${esc(ex ? v.sourceUrl : shared.url)}" style="flex:1;min-width:180px">
        <button class="btn" id="reload" type="button" hidden>↻ Reload from link</button></div>
      <p><button class="btn primary" type="submit">${ex ? 'Save changes' : 'Save recipe'}</button>
        <a class="btn cancel" href="${cancelHref}">Cancel</a></p>
    </form>`;
  const f = $app.querySelector('#f'), msg = $app.querySelector('#impmsg'), impBtn = $app.querySelector('#imp');
  // Text boxes grow to fit what's in them (no drag handle). growAll() is also called after anything fills them in.
  const fit = t => { t.style.height = 'auto'; t.style.height = (t.scrollHeight + 2) + 'px'; };
  const growAll = () => f.querySelectorAll('textarea').forEach(fit);
  f.addEventListener('input', e => { if (e.target.tagName === 'TEXTAREA') fit(e.target); });
  window.onresize = growAll;
  requestAnimationFrame(growAll);
  const saveLabel = ex ? 'Save changes' : 'Save recipe';
  // Cancel asks first only when there's something typed that would be lost.
  const dirty = () => (ex || intake) ? false : ['title', 'ingredients', 'steps', 'summary', 'notes'].some(k => f.elements[k].value.trim());
  $app.querySelectorAll('a.cancel').forEach(a => a.addEventListener('click', e => {
    if (dirty() && !confirm('Discard this recipe? Nothing has been saved yet.')) e.preventDefault();
  }));
  let applyScanned; // defined below, shared by file upload / Drive picker / pasted Drive links
  const runImport = async () => {
    const url = $app.querySelector('#iu').value.trim();
    if (!url) return;
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — import needs it.'; return; }
    if (isGooglePage(url)) { msg.innerHTML = 'That’s a Google results page, not the recipe itself, so there’s nothing to read. Tap the recipe’s source link in the Google result, then share that page instead.<br><a class="btn" href="#/pantry" style="margin-top:8px">← Back to Pantry</a>'; return; }
    msg.textContent = 'Reading the recipe…'; impBtn.disabled = true;
    const showDup = d => { msg.innerHTML = `You already have this one: <a href="#/r/${esc(d.id)}">${esc(d.title)}</a>`; };
    try {
      let dup = await findDup(url, '');
      if (dup) return showDup(dup.r);                       // same page: don't import twice
      const resp = await api({ action: 'import', url });
      if (resp.text) return await applyScanned(resp);       // it was a Drive / Docs link: handled like an uploaded file
      const r = resp.recipe;
      dup = await findDup(r.sourceUrl || url, r.title);
      if (dup && dup.by === 'url') return showDup(dup.r);
      if (dup && !confirm(`You already have a recipe called “${dup.r.title}”. Is this a different one? OK = import it anyway, Cancel = skip.`)) return showDup(dup.r);
      const found = r.ingredients && r.ingredients.length;
      const set = (k, v) => { if (v != null && v !== '') f.elements[k].value = v; };
      set('title', r.title); set('sourceName', r.sourceName); set('sourceUrl', r.sourceUrl || url);
      set('servings', r.servings); set('image', r.image); set('summary', r.summary);
      set('prepTime', r.prepTime); set('cookTime', r.cookTime); set('totalTime', r.totalTime); set('nutrition', r.nutrition);
      if (r.image) showPrev(r.image);
      if (r.category) setTags(String(r.category).split(','));
      set('ingredients', (r.ingredients || []).join('\n')); set('steps', (r.steps || []).join('\n'));
      growAll();
      msg.textContent = found ? 'Imported — check it over, then save.'
        : 'That page has no recipe data I can read. I filled in what I could — add the rest by hand.';
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t import: ' + e.message;
    } finally { impBtn.disabled = false; }
  };
  impBtn.onclick = runImport;
  if (shared.url) runImport();

  // ----- reload a recipe from its web link (only offered for real web pages, not saved scan files) -----
  const reloadBtn = $app.querySelector('#reload'), srcEl = f.elements.sourceUrl;
  const webLink = () => {
    try { const u = new URL(srcEl.value.trim()); return /^https?:$/.test(u.protocol) && !/(^|\.)(drive|docs)\.google\.com$/.test(u.hostname) ? u.href : ''; }
    catch (e) { return ''; }
  };
  const syncReload = () => { reloadBtn.hidden = !webLink(); };
  srcEl.addEventListener('input', syncReload); syncReload();
  reloadBtn.onclick = async () => {
    const url = webLink(); if (!url) return;
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — reloading needs it.'; return; }
    if (!confirm('Reload this recipe from its web link?\n\nThis replaces the title, ingredients, steps, times and other details in this form with a fresh copy from the website. Anything you changed here will be lost. Nothing is saved until you press Save.')) return;
    msg.textContent = 'Reloading from the link…'; reloadBtn.disabled = true;
    try {
      const r = (await api({ action: 'import', url })).recipe;
      if (!(r.ingredients && r.ingredients.length)) { msg.textContent = 'That page has no recipe data I can read, so nothing was changed.'; return; }
      const set = (k, v) => { if (v != null && v !== '') f.elements[k].value = v; };
      set('title', r.title); set('sourceName', r.sourceName); set('servings', r.servings); set('image', r.image); set('summary', r.summary);
      set('prepTime', r.prepTime); set('cookTime', r.cookTime); set('totalTime', r.totalTime); set('nutrition', r.nutrition);
      if (r.image) showPrev(r.image);
      // (reloading from the link leaves your categories alone)
      set('ingredients', r.ingredients.join('\n')); set('steps', (r.steps || []).join('\n'));
      growAll();
      msg.textContent = 'Reloaded from the link — check it over, then save.';
      window.scrollTo(0, 0);
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t reload: ' + e.message;
    } finally { reloadBtn.disabled = false; }
  };

  // ----- new category, recipe picture -----
  const tagBox = $app.querySelector('#tagchips');
  const drawTags = () => {
    f.elements.category.value = [...picked].join(', ');
    tagBox.innerHTML = cats.map(c => `<button type="button" class="chip ${picked.has(c) ? 'on' : ''}" data-tag="${esc(c)}">${esc(c)}</button>`).join('');
  };
  tagBox.onclick = e => { const b = e.target.closest('[data-tag]'); if (!b) return; const t = b.dataset.tag; picked.has(t) ? picked.delete(t) : picked.add(t); drawTags(); };
  const addTag = name => {            // adds (and selects) a tag, matching an existing one regardless of capitals
    const n = String(name || '').replace(/,/g, ' ').trim().slice(0, 30); if (!n) return;
    const have = cats.find(c => c.toLowerCase() === n.toLowerCase());
    const t = have || n; if (!have) cats.push(t); picked.add(t); drawTags();
  };
  const setTags = names => { picked.clear(); names.forEach(addTag); drawTags(); };   // used when an import guesses a category
  drawTags();
  $app.querySelector('#newcat').onclick = () => addTag(prompt('New category name:'));
  const prev = $app.querySelector('#prev');
  const showPrev = u => { if (safeUrl(u)) { prev.referrerPolicy = 'no-referrer'; prev.src = safeUrl(u); prev.hidden = false; } else prev.hidden = true; };
  f.elements.image.oninput = () => showPrev(f.elements.image.value.trim());
  if (ex && ex.image) showPrev(ex.image);
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
  // Fills the form from a file-based result (upload, Drive inbox pick, or pasted Drive link).
  applyScanned = async (d, note) => {
    const r = d.recipe;
    const dup = await findDup('', r.title);
    if (dup && !confirm(`You already have a recipe called “${dup.r.title}”. Is this a different one? OK = fill the form anyway, Cancel = skip.`)) {
      msg.innerHTML = `You already have: <a href="#/r/${esc(dup.r.id)}">${esc(dup.r.title)}</a>`; return;
    }
    const set = (k, v) => { if (v != null && v !== '') f.elements[k].value = v; };
    set('title', r.title); set('servings', r.servings); set('summary', r.summary);
    set('prepTime', r.prepTime); set('cookTime', r.cookTime); set('totalTime', r.totalTime); set('nutrition', r.nutrition);
    if (r.category) setTags(String(r.category).split(','));
    set('ingredients', (r.ingredients || []).join('\n')); set('steps', (r.steps || []).join('\n'));
    // A file or shared text has no web page: mark it as an import, and link the saved original if there is one.
    set('sourceName', 'Import');
    if (d.fileUrl) set('sourceUrl', d.fileUrl);
    $app.querySelector('#rawtxt').textContent = d.text || ''; $app.querySelector('#raw').hidden = !d.text;
    growAll();
    msg.textContent = note || 'Read the file — this is rarely perfect, so check every line before saving.';
    window.scrollTo(0, 0);
  };
  // Opened from the Review tab: fill the form from the queued email.
  if (intake) applyScanned(intake.data, 'From an email in your Review queue — clean it up, then save to post it.');
  // Text shared from another app (e.g. a recipe selected in Gmail): read it straight away.
  if (shared.text) {
    if (!Cfg.url) msg.textContent = 'Connect the Google Sheet first — reading shared text needs it.';
    else {
      msg.textContent = 'Reading the shared text…';
      api({ action: 'parse', text: shared.text }).then(d => applyScanned(d)).catch(e => {
        if (e instanceof AuthError) return settings('Enter the shared passcode, then share the text again.');
        msg.textContent = 'Couldn’t read the shared text: ' + e.message;
      });
    }
  }
  const runScan = async file => {
    if (!file) return;
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — scanning needs it.'; return; }
    msg.textContent = 'Scanning… this can take 10–30 seconds.';
    try {
      const p = await filePayload(file);
      await applyScanned(await api({ action: 'scan', mime: p.mime, data: p.data }));
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t scan: ' + e.message;
    }
  };
  // ----- pick a recipe file from the "Recipe Inbox" folder in Drive -----
  const drvBox = $app.querySelector('#drvbox');
  $app.querySelector('#drv').onclick = async () => {
    if (!Cfg.url) { msg.textContent = 'Connect the Google Sheet first — Drive import needs it.'; return; }
    if (!drvBox.hidden) { drvBox.hidden = true; return; }
    drvBox.hidden = false; drvBox.textContent = 'Looking in your Recipe Inbox folder…';
    try {
      const d = await api({ action: 'drivelist' });
      drvBox.innerHTML = `<p class="muted" style="margin-top:0">Recipe files in your Drive folder “${esc(d.folderName)}”. <a href="${esc(d.folderUrl)}" target="_blank" rel="noopener">Open the folder</a> to add more.</p>` +
        (d.files.length ? d.files.map(x => `<button class="btn" type="button" data-id="${esc(x.id)}" style="display:block;width:100%;text-align:left;margin:4px 0">📄 ${esc(x.name)} <span class="muted">· ${esc(new Date(x.modified).toLocaleDateString())}</span></button>`).join('')
          : '<p>The folder is empty. Add recipe files to it, then tap Import from Drive again.</p>');
    } catch (e) {
      if (e instanceof AuthError) return settings('Enter the shared passcode first.');
      drvBox.textContent = 'Couldn’t open Drive: ' + e.message;
    }
  };
  drvBox.onclick = async e => {
    const b = e.target.closest('[data-id]'); if (!b) return;
    drvBox.hidden = true; msg.textContent = 'Reading the file…';
    try { await applyScanned(await api({ action: 'driveimport', id: b.dataset.id })); }
    catch (err) {
      if (err instanceof AuthError) return settings('Enter the shared passcode first.');
      msg.textContent = 'Couldn’t import: ' + err.message;
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
    const rec = { id: ex ? ex.id : uid(), title: d.title.trim(), category: d.category, sourceName: d.sourceName.trim(), sourceUrl: d.sourceUrl.trim(),
      servings: +d.servings || 1, summary: d.summary.trim(), image: d.image.trim(), favorite: ex ? !!ex.favorite : false,
      prepTime: d.prepTime.trim(), cookTime: d.cookTime.trim(), totalTime: d.totalTime.trim(), nutrition: d.nutrition.trim(), notes: d.notes.trim(),
      ingredients: d.ingredients.split('\n').map(parseIngredient).filter(Boolean),
      steps: d.steps.split('\n').map(s => s.replace(/^\s*\d{1,2}[.)]\s+/, '').trim()).filter(Boolean) };
    // Two Save buttons (top and bottom) — keep them in step.
    if (!picked.size) { toast('Pick at least one category'); return; }
    const setBtns = (off, label) => e.target.querySelectorAll('button[type=submit]').forEach(b => { b.disabled = off; b.textContent = label; });
    setBtns(true, 'Saving…');
    try {
      const dup = ex ? null : await findDup(rec.sourceUrl, rec.title);
      if (dup && !confirm(`“${dup.r.title}” is already in your recipes. Save another copy anyway?`)) { setBtns(false, saveLabel); return; }
      await Store.save(rec);
      if (intake) {   // posted from the Review queue: take it off the queue and go back for the next one
        api({ action: 'intakemark', id: intake.id, status: 'posted' }).catch(() => {});
        let left = 0; try { left = Math.max(0, (+localStorage.getItem('recipes.badge') || 1) - 1); } catch (e) {}
        setBadge(left); toast('Posted to your recipes'); location.hash = '#/review';
      } else location.hash = '#/r/' + rec.id; }
    catch (err) {
      setBtns(false, saveLabel);
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
  const rest = m[2].match(/^(cups?|tbsp|tsp|tablespoons?|teaspoons?|oz|ounces?|lbs?|pounds?|g|grams?|kg|ml|l|cloves?|cans?|pinch|dash|slices?|pkts?|pkgs?|packages?|packets?|sticks?|bunch(?:es)?|heads?|jars?|bottles?|boxes|box|sprigs?|stalks?)\b\.?\s*(.*)$/i);
  return rest ? { amt, unit: rest[1], item: rest[2] } : { amt, unit: '', item: m[2] };
}

// ---------- Pantry tab: type what you have, find saved recipes that use it, optionally get ideas from the web ----------
const WEAK = new Set(['mix', 'fresh', 'frozen', 'canned', 'dried', 'boneless', 'skinless', 'whole', 'large', 'small', 'medium', 'of', 'the', 'a']);
const stem = w => w.endsWith('ies') ? w.slice(0, -3) + 'y' : w.endsWith('oes') ? w.slice(0, -2) : w.endsWith('s') && !w.endsWith('ss') ? w.slice(0, -1) : w;
const words = s => String(s).toLowerCase().replace(/[^a-z\s]/g, ' ').split(/\s+/).filter(Boolean).map(stem);
const splitTerms = q => q.split(/[,;\n]+/).map(t => t.trim()).filter(Boolean);
// Which of the typed ingredients does this recipe use? "chicken breast" matches "boneless skinless chicken breasts"; filler words like "mix" are optional.
function pantryMatch(terms, r) {
  const ing = (r.ingredients || []).map(i => new Set(words(i.item)));
  const hit = terms.filter(t => {
    const all = words(t), core = all.filter(w => !WEAK.has(w)), need = core.length ? core : all;
    return need.length && ing.some(s => need.every(w => s.has(w)));
  });
  return { hit, miss: terms.filter(t => !hit.includes(t)) };
}
// " · This search used 3 points · 12 used today" from Spoonacular's quota headers (left blank if they weren't sent).
function creditsLine(q) {
  if (!q) return '';
  const bits = [];
  if (q.request) bits.push(`This search used ${q.request} ${q.request === 1 ? 'point' : 'points'}`);
  if (q.used != null) bits.push(`${q.used} used today`);
  if (q.left != null) bits.push(`${q.left} left`);
  return bits.length ? ' · ' + bits.join(' · ') : '';
}
// Spoonacular's free credits reset at midnight UTC — show that as the viewer's local time.
function outOfCredits() {
  const n = new Date(), reset = new Date(Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), n.getUTCDate() + 1));
  const when = reset.toLocaleString(undefined, { weekday: 'short', hour: 'numeric', minute: '2-digit' });
  const back = document.createElement('div'); back.className = 'modal';
  back.innerHTML = `<div class="sheet" role="alertdialog" aria-label="Out of credits"><h2 style="margin-top:0">Out of web-search credits</h2>
    <p>We’ve used up today’s free Spoonacular credits. They reset <b>${esc(when)}</b>.</p><p class="muted">Searching your saved recipes still works.</p>
    <div class="row" style="justify-content:center"><button class="btn primary">OK</button></div></div>`;
  back.onclick = e => { if (e.target === back || e.target.tagName === 'BUTTON') back.remove(); };
  document.body.appendChild(back);
}
async function pantry() {
  // The last search (and its Spoonacular results) is kept on this device, so coming back from a recipe, another app or a
  // failed share finds it still there — and viewing it again doesn't spend any Spoonacular points.
  let mem = {}; try { mem = JSON.parse(localStorage.getItem('recipes.pantry2') || '{}') || {}; } catch (e) {}
  const keep = () => { try { localStorage.setItem('recipes.pantry2', JSON.stringify(mem)); } catch (e) {} };
  const last = mem.q || '';
  $app.innerHTML = `
    <h1>🥘 Pantry</h1>
    <p class="muted">Type what you have, separated by commas — like <i>chicken breast, stuffing mix</i> — and find recipes that use it.</p>
    <form class="row" id="pf"><span class="clearwrap"><input id="pq" placeholder="chicken breast, stuffing mix" value="${esc(last)}"><button type="button" class="clearx" id="pclear" aria-label="Clear search" title="Clear"${last ? '' : ' hidden'}>×</button></span><button class="btn primary">Search</button></form>
    <div id="pres"></div><div id="pweb"></div>`;
  const out = $app.querySelector('#pres'), web = $app.querySelector('#pweb');
  const run = async () => {
    const terms = splitTerms($app.querySelector('#pq').value);
    web.innerHTML = '';
    if (!terms.length) { out.innerHTML = ''; return; }
    const q = terms.join(', ');
    if (mem.q !== q) mem = { q };            // a different search: forget the old web results
    keep();
    out.innerHTML = '<p class="muted">Searching your recipes…</p>';
    const all = await Store.list();
    const found = all.map(r => ({ r, ...pantryMatch(terms, r) })).filter(x => x.hit.length)
      .sort((a, b) => b.hit.length - a.hit.length || (a.r.ingredients.length - a.hit.length) - (b.r.ingredients.length - b.hit.length) || a.r.title.localeCompare(b.r.title));
    out.innerHTML = `<h2>Your saved recipes</h2>` + (found.length ? `<div class="list">${found.map(({ r, hit, miss }) => `
      <a class="lrow" href="#/r/${esc(encodeURIComponent(r.id))}">
        <div class="lthumb" style="${r.image ? `background-image:url('${esc(safeUrl(r.image))}')` : ''}">${r.image ? '' : '🍽️'}</div>
        <div class="lbody"><b>${esc(r.title)}</b>
          <span class="muted">Uses ${hit.length} of ${terms.length} · ${Math.max(0, r.ingredients.length - hit.length)} other ingredients</span>
          ${miss.length ? `<span class="muted" style="display:block">Doesn’t use: ${esc(miss.join(', '))}</span>` : ''}</div></a>`).join('')}</div>`
      : '<p class="muted">No saved recipe uses those ingredients.</p>')
      + `<div class="row"><button class="btn primary" id="webgo">🥄 Search Spoonacular</button>
          <a class="btn alt" id="google"target="_blank" rel="noopener" href="https://www.google.com/search?q=${encodeURIComponent(terms.join(' ') + ' recipe')}">🔍 Search Google</a></div>
        <p class="muted">Found one on Google? Use your phone’s Share button and pick this app to import it.</p>`;
    const showWeb = (d, fresh) => {
        web.innerHTML = '<h2>From Spoonacular</h2>' + (d.results.length ? `<div class="list">${d.results.map((x, n) => `
          <div class="lrow"><div class="lthumb" style="${x.image ? `background-image:url('${esc(safeUrl(x.image))}')` : ''}">${x.image ? '' : '🍽️'}</div>
            <div class="lbody"><b style="white-space:normal">${esc(x.title)}</b>
              <span class="muted" style="display:block">${x.used.length ? 'You have: ' + esc(x.used.join(', ')) : ''}${x.minutes ? ' · ' + x.minutes + ' min' : ''}</span>
              ${x.missed.length ? `<span class="muted" style="display:block">Also needs: ${esc(x.missed.join(', '))}</span>` : ''}
              <span class="row" style="margin:6px 0 0">${safeUrl(x.url) ? `<a class="btn" href="${esc(safeUrl(x.url))}" target="_blank" rel="noopener">View</a><button class="btn primary" data-save="${n}">Save to my recipes</button>` : safeUrl(x.page) ? `<a class="btn" href="${esc(safeUrl(x.page))}" target="_blank" rel="noopener">View</a>` : ''}</span></div></div>`).join('')}</div>`
          : `<p class="muted">Spoonacular has no recipe that uses all of: ${esc(terms.join(', '))}. Try dropping one ingredient, or use Search Google.</p>`)
          + (d.results.length && d.results.length < 5 ? `<p class="muted">Only ${d.results.length} found that use everything you typed — drop an ingredient for more ideas.</p>` : '')
          + `<p class="muted">Recipe ideas powered by <a href="https://spoonacular.com/food-api" target="_blank" rel="noopener">spoonacular</a>.${fresh ? creditsLine(d.quota) : ' · Saved from your last search — no points used.'}</p>`;
        web.querySelectorAll('[data-save]').forEach(s => s.onclick = () => { const x = d.results[+s.dataset.save]; window.__shared = { title: x.title, url: x.url }; location.hash = '#/add'; });
    };
    if (mem.web) showWeb(mem.web, false);
    out.querySelector('#webgo').onclick = async e => {
      const b = e.currentTarget; b.disabled = true; web.innerHTML = '<p class="muted">Searching Spoonacular…</p>';
      try {
        const d = await api({ action: 'websuggest', terms });
        mem.web = d; keep(); showWeb(d, true);
      } catch (err) {
        if (err instanceof AuthError) throw err;
        web.innerHTML = `<p>${esc(err.message)}</p>${err.code === 'quota' ? `<p class="muted">${creditsLine(err.quota).replace(/^ · /, '')}</p>` : ''}`;
        if (err.code === 'quota') outOfCredits();
      }
      b.disabled = false;
    };
  };
  $app.querySelector('#pf').onsubmit = e => { e.preventDefault(); run().catch(err => { out.innerHTML = `<p>Couldn’t search: ${esc(err.message)}</p>`; }); };
  const pq = $app.querySelector('#pq'), px = $app.querySelector('#pclear');
  pq.oninput = () => { px.hidden = !pq.value; };
  px.onclick = () => { pq.value = ''; px.hidden = true; mem = {}; keep(); out.innerHTML = ''; web.innerHTML = ''; pq.focus(); };
  if (last) run().catch(() => {});
}

// Small pop-up month calendar (weeks start Sunday): tap a day and onPick(date) runs. Tap outside or Cancel to close.
function pickDay(onPick) {
  const now = new Date(); let Y = now.getFullYear(), M = now.getMonth() + 1;
  const back = document.createElement('div'); back.className = 'modal';
  document.body.appendChild(back);
  const close = () => back.remove();
  back.onclick = e => { if (e.target === back) close(); };
  const draw = () => {
    const first = new Date(Y, M - 1, 1), days = new Date(Y, M, 0).getDate(), today = ymd(now);
    let cells = '';
    for (let i = 0; i < first.getDay(); i++) cells += '<span></span>';
    for (let d = 1; d <= days; d++) {
      const date = Y + '-' + String(M).padStart(2, '0') + '-' + String(d).padStart(2, '0');
      cells += `<button type="button" class="pd ${date === today ? 'today' : ''}" data-d="${date}">${d}</button>`;
    }
    back.innerHTML = `<div class="sheet" role="dialog" aria-label="Pick a night">
      <div class="row calnav"><button type="button" class="btn" id="pm">‹</button><b>${esc(first.toLocaleDateString(undefined, { month: 'long', year: 'numeric' }))}</b><button type="button" class="btn" id="nm">›</button></div>
      <div class="pdgrid">${['S', 'M', 'T', 'W', 'T', 'F', 'S'].map(x => `<i>${x}</i>`).join('')}${cells}</div>
      <div class="row" style="justify-content:center"><button type="button" class="btn" id="pc">Cancel</button></div></div>`;
    back.querySelector('#pm').onclick = () => { M--; if (M < 1) { M = 12; Y--; } draw(); };
    back.querySelector('#nm').onclick = () => { M++; if (M > 12) { M = 1; Y++; } draw(); };
    back.querySelector('#pc').onclick = close;
    back.querySelectorAll('[data-d]').forEach(b => b.onclick = () => { const d = b.dataset.d; close(); onPick(d); });
  };
  draw();
}

// ---------- Calendar tab: dinners planned by date, weeks start on Sunday ----------
const ymd = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const niceDate = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' }); };
async function calendar(ym) {
  $app.innerHTML = '<p class="muted">Loading the calendar…</p>';
  const [plan, recipes] = await Promise.all([api({ action: 'planlist' }).then(d => d.plan), Store.list().catch(() => [])]);
  const byId = new Map(recipes.map(r => [r.id, r]));
  const now = new Date();
  let [Y, M] = /^\d{4}-\d{2}$/.test(ym || '') ? ym.split('-').map(Number) : [now.getFullYear(), now.getMonth() + 1];
  const first = new Date(Y, M - 1, 1), days = new Date(Y, M, 0).getDate(), lead = first.getDay();
  const prev = new Date(Y, M - 2, 1), next = new Date(Y, M, 1), key = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  const today = ymd(now);
  const cells = [];
  for (let i = 0; i < lead; i++) cells.push('<div class="day blank"></div>');
  for (let d = 1; d <= days; d++) {
    const date = Y + '-' + String(M).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    const items = plan.filter(p => p.date === date);
    cells.push(`<div class="day ${date === today ? 'today' : ''} ${items.length ? 'has' : ''}"><div class="dnum"><span class="dname">${esc(niceDate(date))}</span><span class="dn">${d}</span></div>
      ${items.map(p => { const rec = byId.get(p.recipeId); const title = rec ? rec.title : p.title;
        return `<div class="plan">${rec ? `<a href="#/r/${esc(encodeURIComponent(p.recipeId))}">${esc(title)}</a>` : `<span title="This recipe was deleted">${esc(title)}</span>`}<button class="x" data-del="${esc(p.id)}" title="Remove from this night" aria-label="Remove">✕</button></div>`; }).join('')}</div>`);
  }
  const title = first.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
  $app.innerHTML = `
    <h1>📅 Dinner calendar</h1>
    <div class="row calnav"><a class="btn" href="#/calendar/${key(prev)}">‹</a><b>${esc(title)}</b><a class="btn" href="#/calendar/${key(next)}">›</a>
      <a class="btn" href="#/calendar">Today</a></div>
    ${plan.length ? '' : '<p class="muted">Nothing planned yet. Open a recipe, pick a date, and tap “Add to calendar”.</p>'}
    <div class="calgrid"><div class="dow">Sun</div><div class="dow">Mon</div><div class="dow">Tue</div><div class="dow">Wed</div><div class="dow">Thu</div><div class="dow">Fri</div><div class="dow">Sat</div>${cells.join('')}</div>`;
  $app.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
    if (!confirm('Remove this recipe from that night?')) return;
    try { await api({ action: 'plandelete', id: b.dataset.del }); calendar(key(first)); } catch (err) { toast('Couldn’t remove: ' + err.message); }
  });
}

// ---------- Shopping tab: one weekly list shared by everyone ----------
const shopText = i => ((i.amt == null ? '' : fmtAmt(i.amt)) + ' ' + (i.unit || '') + ' ' + i.item).replace(/\s+/g, ' ').trim();
async function shopping() {
  $app.innerHTML = '<p class="muted">Loading the shopping list…</p>';
  let items = (await api({ action: 'shoplist' })).items;
  const draw = () => {
    const open = items.filter(i => !i.done), done = items.filter(i => i.done);
    const row = i => `<li class="shop ${i.done ? 'done' : ''}"><label><input type="checkbox" data-id="${esc(i.id)}" ${i.done ? 'checked' : ''}>
      <span><b>${esc(shopText(i))}</b>${i.recipes.length ? `<span class="muted"> · ${esc(i.recipes.join(', '))}</span>` : ''}</span></label>
      <button class="x" data-del="${esc(i.id)}" title="Remove" aria-label="Remove">✕</button></li>`;
    $app.innerHTML = `
      <h1>🛒 Shopping</h1>
      <form class="row" id="addf"><input id="newitem" placeholder="Add an item (e.g. 2 lb potatoes)" style="flex:1;min-width:0"><button class="btn primary">Add</button></form>
      <div class="row tools"><button class="btn" id="share">📤 Send to Keep / copy</button><button class="btn" id="refresh">↻ Refresh</button></div>
      ${items.length ? '' : '<p class="muted">Nothing on the list yet. Open a recipe, tick what you need, and tap “Add to weekly list”.</p>'}
      <ul class="shoplist">${open.map(row).join('')}</ul>
      ${done.length ? `<h2>In the cart</h2><ul class="shoplist">${done.map(row).join('')}</ul>
        <div class="row"><button class="btn" id="cleardone">Clear ticked items</button></div>` : ''}
      ${items.length ? '<div class="row"><button class="btn" id="clearall">Clear the whole list</button></div>' : ''}`;
    const q = s => $app.querySelector(s);
    q('#addf').onsubmit = async e => {
      e.preventDefault();
      const v = q('#newitem').value.trim(); if (!v) return;
      const p = parseIngredient(v);
      try { items = (await api({ action: 'shopadd', items: [{ amt: p.amt, unit: p.unit, item: p.item }], recipe: '' })).items; draw(); }
      catch (err) { toast('Couldn’t add: ' + err.message); }
    };
    $app.querySelectorAll('input[type=checkbox][data-id]').forEach(c => c.onchange = async () => {
      const it = items.find(x => x.id === c.dataset.id); if (!it) return;
      it.done = c.checked; draw();
      try { await api({ action: 'shopset', id: it.id, done: it.done }); } catch (err) { it.done = !it.done; draw(); toast('Couldn’t save that tick: ' + err.message); }
    });
    $app.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
      const id = b.dataset.del; const was = items; items = items.filter(x => x.id !== id); draw();
      try { await api({ action: 'shopdelete', mode: 'one', id }); } catch (err) { items = was; draw(); toast('Couldn’t remove: ' + err.message); }
    });
    q('#refresh').onclick = async () => { try { items = (await api({ action: 'shoplist' })).items; draw(); toast('Updated'); } catch (err) { toast(err.message); } };
    const clear = mode => async () => {
      if (!confirm(mode === 'all' ? 'Clear the whole shopping list for everyone?' : 'Remove the ticked items for everyone?')) return;
      try { await api({ action: 'shopdelete', mode }); items = mode === 'all' ? [] : items.filter(x => !x.done); draw(); }
      catch (err) { toast('Couldn’t clear: ' + err.message); }
    };
    if (q('#cleardone')) q('#cleardone').onclick = clear('done');
    if (q('#clearall')) q('#clearall').onclick = clear('all');
    q('#share').onclick = async () => {
      const list = items.filter(i => !i.done); if (!list.length) return toast('Nothing left to buy');
      const text = list.map(i => '☐ ' + shopText(i)).join('\n');
      const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Macintosh/i.test(navigator.userAgent));
      if (isMobile && navigator.share) { try { await navigator.share({ title: 'Shopping list', text }); return; } catch (e) { if (e.name === 'AbortError') return; } }
      try { await navigator.clipboard.writeText(text); toast('List copied — paste it into an email or note'); } catch (e) { prompt('Copy this list:', text); }
    };
  };
  draw();
}

// ---------- router ----------
async function route() {
  const h = location.hash || '#/';
  window.scrollTo(0, 0);
  try {
    document.querySelectorAll('.tabs a').forEach(a => a.classList.toggle('on', a.dataset.t === (h.startsWith('#/review') ? 'review' : h.startsWith('#/shop') ? 'shop' : h.startsWith('#/pantry') ? 'pantry' : h.startsWith('#/calendar') ? 'cal' : h.startsWith('#/add') || h.startsWith('#/edit') ? '' : 'home')));
    if (h.startsWith('#/review')) return await review();
    if (h.startsWith('#/shop')) return await shopping();
    if (h.startsWith('#/pantry')) return await pantry();
    if (h.startsWith('#/calendar')) return await calendar(h.slice(11));
    if (h.startsWith('#/settings')) return settings();
    if (h.startsWith('#/r/')) return await recipe(decodeURIComponent(h.slice(4)));
    if (h.startsWith('#/edit/')) {
      const ex = await Store.get(decodeURIComponent(h.slice(7)));
      return ex ? await add(null, ex) : ($app.innerHTML = '<p>Recipe not found. <a href="#/">Back</a></p>');
    }
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
    const text = (p.get('text') || '').trim();
    let url = p.get('url') || '';
    if (!url) { const m = text.match(/https?:\/\/\S+/); url = m ? m[0] : ''; }
    // A shared web page is "title + link". Shared *text* (e.g. a recipe selected in an email) has real content
    // beyond any link: several lines or a few hundred characters.
    const body = text.replace(/https?:\/\/\S+/g, '').trim();
    const isRecipeText = body.length >= 150 || body.split('\n').filter(l => l.trim()).length >= 4;
    window.__shared = isRecipeText ? { title: p.get('title') || '', url: '', text: body } : { title: p.get('title') || '', url };
    history.replaceState(null, '', location.pathname + '#/add');
  }
})();

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
try { const b = +localStorage.getItem('recipes.badge'); if (b) setBadge(b); } catch (e) {}
route();
refreshBadge();
