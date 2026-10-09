/* L2 Reborn Recipe Tracker - web edition (reference tabs). Data comes from data.js (exported from the workbook). */
(function () {
'use strict';
const D = window.L2_DATA;
const GRADES = ['D', 'C', 'B', 'A', 'S', 'S80', 'S84'];
const CHRON = ['C1', 'C2', 'C3', 'C4', 'C5', 'IL', 'K', 'HB', 'GF', 'F'];
const $ = (s, r) => (r || document).querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const lc = s => String(s == null ? '' : s).toLowerCase();
const num = n => (typeof n === 'number' ? n.toLocaleString('en-US', { maximumFractionDigits: 3 }) : esc(n));
const pct = v => (typeof v === 'number' ? Math.round(v * 1000) / 10 + '%' : esc(v));
function rate(v) {
  if (typeof v !== 'number') return esc(v == null ? '' : v);
  let s = (v * 100).toFixed(4).replace(/0+$/, '');
  if (s.endsWith('.')) s += '0';
  return s + '%';
}
const rateSort = v => (typeof v === 'number' ? v : -1);
const uniq = a => [...new Set(a)];
const byOrder = order => (a, b) => (order.indexOf(a) < 0 ? 99 : order.indexOf(a)) - (order.indexOf(b) < 0 ? 99 : order.indexOf(b));

/* ------------------------------------------------------------------ indexes */
const matByName = {}; D.materials.forEach(m => (matByName[m.name] = m));
const itemByName = {}; D.equipment.forEach(i => (itemByName[i.name] = i));
const rawCols = D.rawBreakdown.cols;
const rawByMat = {}; D.rawBreakdown.rows.forEach(([n, v]) => (rawByMat[n] = v));
const catByName = {}; D.catalogue.forEach(([cat, n, g, intro]) => (catByName[n] = { cat, g, intro }));
const esByKey = {}; D.equipSourcing.rows.forEach(r => (esByKey[r[0]] = esByKey[r[0]] || []).push(r));
const msByKey = {}; D.matSourcing.rows.forEach(r => (msByKey[r[0]] = msByKey[r[0]] || []).push(r));
const usedInItems = {}; D.equipment.forEach(it => it.ing.forEach(([n]) => (usedInItems[n] = usedInItems[n] || []).push(it)));
const usedInMats = {}; D.materials.forEach(m => m.ing.forEach(([n]) => (usedInMats[n] = usedInMats[n] || []).push(m)));
const recipeOf = {}; D.equipment.forEach(it => it.rk && (recipeOf[it.rk] = it)); D.materials.forEach(m => (recipeOf['Recipe: ' + m.name] = m));
const E = window.L2Engine(D);
D.equipment.forEach(it => { const sn = E.scrollName(it); if (sn) recipeOf[sn] = it; });

/* warehouse data: lives in the shared database (or in this browser in test mode) - see backend.js */
const B = window.L2Backend;
let store = { stock: {}, learned: {}, meta: {} };
let session = null, loaded = false, loadErr = null, loading = null, unsub = null, onData = null;
const short = e => String(e || '').split('@')[0];
/* Friendly name: config.js can list exact names (L2_CONFIG.names); otherwise the first word of the e-mail name, capitalised. */
const nice = e => { const h = short(e).toLowerCase(), m = (window.L2_CONFIG || {}).names || {}; if (m[h]) return m[h]; const w = h.split(/[^a-z]+/)[0]; return w ? w[0].toUpperCase() + w.slice(1) : h; };
const when = iso => (iso ? new Date(iso).toLocaleString('en-GB', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }) : '');
const sgn = n => (n > 0 ? '+' : n < 0 ? '−' : '') + num(Math.abs(n));
const kindCls = k => ({ 'Raw material': 't-raw', 'Equipment part': 't-part', 'Crystal / Gemstone': 't-crystal', 'Recipe scroll': 't-recipe' }[k] || '');

function kindOf(name) {
  if (itemByName[name]) return { label: 'Equipment · ' + itemByName[name].grade, cls: 'g-' + itemByName[name].grade };
  if (matByName[name]) return { label: 'Material · Tier ' + matByName[name].tier, cls: 't-' + matByName[name].tier };
  if (rawCols.includes(name)) return { label: 'Raw material', cls: 't-raw' };
  const c = catByName[name];
  if (c) {
    if (c.cat.startsWith('Crystal')) return { label: 'Crystal / Gemstone', cls: 't-crystal' };
    if (c.cat === 'Equipment Part') return { label: 'Equipment part', cls: 't-part' };
    if (c.cat === 'Recipe') return { label: 'Recipe scroll', cls: 't-recipe' };
  }
  if (/^Recipe: /.test(name)) return { label: 'Recipe scroll', cls: 't-recipe' };
  if (esByKey[name]) return { label: 'Equipment part', cls: 't-part' };
  if (/^(Crystal:|Gemstone)/.test(name)) return { label: 'Crystal / Gemstone', cls: 't-crystal' };
  return { label: 'Item', cls: '' };
}
const gradeColor = g => `var(--g-${g})`;
const tierColor = t => `var(--t-${t})`;
const chip = (text, cls) => `<span class="chip ${cls || ''}">${esc(text)}</span>`;
const link = (name, extra) => `<button class="lnk ${extra || ''}" data-open="${esc(name)}">${esc(name)}</button>`;
const isCurrency = n => /^(Crystal:|Gemstone)/.test(n);
const ingList = ing => ing.map(([n, q]) => `<div class="ing"><span class="q">${num(q)}×</span>${link(n, isCurrency(n) ? 'dim' : '')}</div>`).join('');

/* --------------------------------------------------------------- table engine */
function makeTable(host, cfg, st) {
  st.q = st.q || ''; st.f = st.f || {}; st.sort = st.sort || null; st.t = st.t || {};
  cfg.toggles && cfg.toggles.forEach(t => { if (!(t.id in st.t)) st.t[t.id] = !!t.def; });
  st.limit = st.limit || cfg.pageSize || 150;
  const hay = cfg.rows.map(r => lc(cfg.search(r)));
  const visibleCols = () => cfg.cols.filter(c => !c.toggle || st.t[c.toggle]);
  host.innerHTML = `
    <div class="controls">
      <input type="search" placeholder="${esc(cfg.placeholder || 'Search…')}" aria-label="Search this table" value="${esc(st.q)}">
      ${(cfg.filters || []).map(f => `<select data-f="${f.id}" aria-label="${esc(f.label)}"><option value="">${esc(f.label)}: all</option>${f.options.map(o => `<option value="${esc(o)}"${st.f[f.id] === o ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select>`).join('')}
      ${(cfg.toggles || []).map(t => `<label class="tog"><input type="checkbox" data-t="${t.id}"${st.t[t.id] ? ' checked' : ''}> ${esc(t.label)}</label>`).join('')}
      <button class="linkbtn" data-reset hidden>Clear filters</button>
      <span class="count"></span>
    </div>
    <div class="tablewrap"><table><thead></thead><tbody></tbody></table><div class="more"></div></div>`;
  const thead = $('thead', host), tbody = $('tbody', host), more = $('.more', host), count = $('.count', host), reset = $('[data-reset]', host);

  function filtered() {
    const terms = lc(st.q).split(/\s+/).filter(Boolean);
    let idx = [];
    for (let i = 0; i < cfg.rows.length; i++) {
      const r = cfg.rows[i];
      let ok = true;
      for (const f of cfg.filters || []) { const v = st.f[f.id]; if (v && !f.test(r, v)) { ok = false; break; } }
      if (ok) for (const t of terms) if (!hay[i].includes(t)) { ok = false; break; }
      if (ok) idx.push(i);
    }
    if (st.sort) {
      const c = cfg.cols.find(x => x.key === st.sort.key), dir = st.sort.dir;
      const key = c.sort || c.get;
      idx.sort((a, b) => {
        let x = key(cfg.rows[a]), y = key(cfg.rows[b]);
        if (x == null) x = ''; if (y == null) y = '';
        const d = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
        return (d || a - b) * dir;
      });
    }
    return idx.map(i => cfg.rows[i]);
  }
  function drawHead() {
    thead.innerHTML = '<tr>' + visibleCols().map(c => {
      const on = st.sort && st.sort.key === c.key;
      return `<th class="${c.cls || ''}">${c.nosort ? esc(c.label) : `<button data-sort="${c.key}" aria-label="Sort by ${esc(c.label)}">${esc(c.label)}${on ? `<span class="arr">${st.sort.dir > 0 ? '▲' : '▼'}</span>` : ''}</button>`}</th>`;
    }).join('') + '</tr>';
  }
  function draw() {
    const rows = filtered(), cols = visibleCols();
    const shown = rows.slice(0, st.limit);
    let prevKey = null;
    tbody.innerHTML = shown.map(r => {
      const rc = cfg.stripe ? cfg.stripe(r) : '';
      const key = cfg.group ? cfg.group(r) : null;
      const cont = key !== null && key === prevKey;
      prevKey = key;
      return `<tr class="${cont ? 'cont' : cfg.group ? 'grp' : ''}">${cols.map((c, i) => `<td class="${c.cls || ''}"${i === 0 && rc ? ` style="--rc:${rc}"` : ''}>${cont && (cfg.groupCols || []).includes(c.key) ? '' : c.html(r)}</td>`).join('')}</tr>`;
    }).join('') || `<tr><td class="empty" colspan="${cols.length}">Nothing matches these filters.</td></tr>`;
    const left = rows.length - shown.length;
    more.innerHTML = left > 0 ? `Showing ${shown.length} of ${rows.length} · <button class="linkbtn" data-more>show ${Math.min(left, cfg.pageSize || 150)} more</button> · <button class="linkbtn" data-all>show all</button>` : '';
    count.textContent = `${rows.length} of ${cfg.rows.length} ${cfg.noun || 'rows'}`;
    const active = st.q || Object.values(st.f).some(Boolean);
    reset.hidden = !active;
  }
  drawHead(); draw();
  host.addEventListener('input', e => {
    if (e.target.matches('input[type=search]')) { st.q = e.target.value; st.limit = cfg.pageSize || 150; draw(); }
  });
  host.addEventListener('change', e => {
    const f = e.target.dataset.f, t = e.target.dataset.t;
    if (f) { st.f[f] = e.target.value; st.limit = cfg.pageSize || 150; draw(); }
    if (t) { st.t[t] = e.target.checked; drawHead(); draw(); }
  });
  host.addEventListener('click', e => {
    const s = e.target.closest('[data-sort]');
    if (s) {
      const k = s.dataset.sort;
      if (!st.sort || st.sort.key !== k) st.sort = { key: k, dir: 1 };
      else if (st.sort.dir === 1) st.sort.dir = -1; else st.sort = null;
      drawHead(); draw(); return;
    }
    if (e.target.closest('[data-more]')) { st.limit += cfg.pageSize || 150; draw(); }
    if (e.target.closest('[data-all]')) { st.limit = 1e9; draw(); }
    if (e.target.closest('[data-reset]')) {
      st.q = ''; st.f = {}; $('input[type=search]', host).value = '';
      host.querySelectorAll('select').forEach(s => (s.value = '')); draw();
    }
  });
}

/* --------------------------------------------------------------------- views */
const state = {};
const st = id => (state[id] = state[id] || {});
const head = (title, lead) => `<h1>${esc(title)}</h1><div class="orn"><i></i><b></b><i></i></div>${lead ? `<p class="lead">${esc(lead)}</p>` : ''}`;
const gradeOpts = GRADES.filter(g => D.equipment.some(i => i.grade === g));
const chronOpts = CHRON.filter(c => D.equipment.some(i => i.intro === c) || D.equipSourcing.rows.some(r => r[3] === c) || D.matSourcing.rows.some(r => r[1] === c));

function startView(el) {
  const L = D.start, sections = []; let cur = null;
  L.slice(2).forEach(l => {
    if (l.b) { cur = { title: l.t, lines: [] }; sections.push(cur); } else if (cur) cur.lines.push(l);
  });
  const sec = p => sections.find(s => s.title.startsWith(p));
  const tips = (sec('TIPS') || { lines: [] }).lines.filter(l => !/filter buttons|repeat on every row/i.test(l.t));
  const swatches = s => (s ? s.lines.map(l => `<div><span class="sw" style="background:#${l.fill || 'ccc'}"></span>${esc(l.t)}</div>`).join('') : '');
  el.innerHTML = head(L[0].t, L[1].t) + `
    <div class="banner"><b>Web edition — shared warehouse.</b> The reference tabs are read from the workbook (${esc(D.meta.version)}). Warehouse, Requirements Check and History are the clan tools: sign in and both of you see and edit the same stock.</div>
    <div class="cols2">
      <div class="card"><h2>How to use this site</h2>
        <ul>
          <li>Use the search box at the top to jump to any item, material, recipe scroll or part.</li>
          <li>Click any underlined name in a table to open its detail panel: ingredients, raw-material totals, where to get it and where it is used.</li>
          <li>Every table has a search box and filters, and column headers sort on click.</li>
          <li>Rate columns show the chance (%) that the monster drops or spoils the scroll or part.</li>
        </ul></div>
      <div class="card"><h2>Tabs</h2>
        <ul>
          <li><b>Material Recipes</b> — what each Tier 1–4 material is made from.</li>
          <li><b>Raw Material Breakdown</b> — each material reduced to raw materials, per unit.</li>
          <li><b>Material Recipe Sourcing</b> — where to get each material's recipe scroll.</li>
          <li><b>Equipment Recipes</b> — every weapon, armor and jewelry recipe (D to S84).</li>
          <li><b>Equipment Sourcing</b> — where to get each equipment recipe scroll and part.</li>
          <li><b>Notes</b> and <b>Audit Log</b> — methodology, open data issues and the per-row check notes.</li>
        </ul></div>
    </div>
    <div class="cols2">
      <div class="card"><h2>Scrolls &amp; currency</h2>${tips.map(l => `<p>${esc(l.t)}</p>`).join('')}</div>
      <div class="card"><h2>Chronicle labels</h2>${(sec('CHRONICLE') || { lines: [] }).lines.map(l => `<p>${esc(l.t)}</p>`).join('')}
        <h2 style="margin-top:14px">Source types</h2>${(sec('SOURCE') || { lines: [] }).lines.map(l => `<p>${esc(l.t)}</p>`).join('')}</div>
    </div>
    <div class="cols2">
      <div class="card"><h2>Colour key — materials</h2>${swatches(sections.find(s => /COLOR KEY — Materials/.test(s.title)))}</div>
      <div class="card"><h2>Colour key — equipment grade</h2>${swatches(sections.find(s => /COLOR KEY — Equipment/.test(s.title)))}</div>
    </div>`;
}

function materialsView(el) {
  /* "Recipe learned" tick: a shared reminder (stored in the same learned list as the equipment recipes).
     It does not affect the Requirements Check or the warehouse. Needs a sign-in, like the clan tools. */
  const can = B.mode === 'local' || !!session;
  el.innerHTML = head('Material Recipes', 'What each Tier 1–4 material is made from. Tier 1 uses raw materials only; higher tiers build on lower ones.') +
    (can ? '' : '<p class="hint">Sign in (Warehouse tab) to tick which material recipes you have already learned.</p>') + '<div id="t"></div>';
  const cv = current;
  if (can && !loaded && !loadErr) ensureLoaded().then(() => { if (current === cv && $('#view') === el) rerender(); });
  onData = () => el.querySelectorAll('[data-ml]').forEach(c => { c.checked = !!store.learned[c.dataset.ml]; });
  makeTable($('#t', el), {
    rows: D.materials, noun: 'materials', placeholder: 'Search material or ingredient…',
    search: m => m.name + ' ' + m.ing.map(i => i[0]).join(' '),
    stripe: m => tierColor(m.tier),
    filters: [
      { id: 'tier', label: 'Tier', options: ['1', '2', '3', '4'], test: (m, v) => String(m.tier) === v },
      { id: 'intro', label: 'Introduced', options: uniq(D.materials.map(m => m.intro)).sort(byOrder(CHRON)), test: (m, v) => m.intro === v },
    ].concat(can ? [{ id: 'rl', label: 'Recipe', options: ['Learned', 'Not learned yet'], test: (m, v) => !!store.learned[m.name] === (v === 'Learned') }] : []),
    cols: [
      { key: 'name', label: 'Material', get: m => m.name, html: m => link(m.name) },
      ...(can ? [{ key: 'rl', label: 'Recipe learned', get: m => (store.learned[m.name] ? 0 : 1),
        html: m => `<label class="tog mlk"><input type="checkbox" data-ml="${esc(m.name)}"${store.learned[m.name] ? ' checked' : ''}${loaded ? '' : ' disabled'} aria-label="Recipe for ${esc(m.name)} learned"></label>` }] : []),
      { key: 'tier', label: 'Tier', get: m => m.tier, html: m => chip('Tier ' + m.tier, 't-' + m.tier) },
      { key: 'intro', label: 'Introduced', get: m => CHRON.indexOf(m.intro), html: m => `<span class="chr">${esc(m.intro)}</span>` },
      { key: 'out', label: 'Output qty', cls: 'num', get: m => m.out, html: m => num(m.out) },
      { key: 'ing', label: 'Ingredients', nosort: true, get: m => '', html: m => ingList(m.ing) },
    ],
  }, st('materials'));
  $('#t', el).addEventListener('change', async e => {
    const c = e.target.closest('[data-ml]'); if (!c) return;
    const item = c.dataset.ml, chk = c.checked, prev = !!store.learned[item];
    if (chk) store.learned[item] = true; else delete store.learned[item];
    try { await B.setLearned(item, chk); } catch (err) { if (prev) store.learned[item] = true; else delete store.learned[item]; c.checked = prev; toast('Could not save — ' + errText(err)); }
  });
}

function rawView(el) {
  const R = D.rawBreakdown, s = st('raw'); s.q = s.q || ''; if (s.hide === undefined) s.hide = false;
  el.innerHTML = head('Raw Material Breakdown', R.title) + `
    <div class="controls"><input type="search" placeholder="Search material…" value="${esc(s.q)}" aria-label="Search materials">
      <label class="tog"><input type="checkbox" id="hz"${s.hide ? ' checked' : ''}> Hide raw materials no recipe uses yet</label><span class="count"></span></div>
    <div class="tablewrap"><table class="matrix"><thead></thead><tbody></tbody></table></div>`;
  function draw() {
    const used = R.cols.map((c, i) => R.rows.some(r => r[1][i] > 0));
    const cols = R.cols.map((c, i) => i).filter(i => !s.hide || used[i]);
    $('thead', el).innerHTML = '<tr><th>Material</th>' + cols.map(i => `<th class="col">${link(R.cols[i], 'dim')}</th>`).join('') + '</tr>';
    const q = lc(s.q);
    const rows = R.rows.filter(r => !q || lc(r[0]).includes(q));
    $('tbody', el).innerHTML = rows.map(([n, v]) => `<tr><td style="--rc:${matByName[n] ? tierColor(matByName[n].tier) : 'transparent'}">${link(n)}</td>` +
      cols.map(i => `<td class="v ${v[i] ? 'nz' : 'z'}">${v[i] ? num(v[i]) : '·'}</td>`).join('') + '</tr>').join('');
    $('.count', el).textContent = `${rows.length} of ${R.rows.length} materials`;
    fit();
  }
  // Wide screens: horizontal (wrapped) names. Narrow screens: vertical names so every column still fits.
  function fit() {
    const wrap = $('.tablewrap', el), tb = $('table.matrix', el); if (!wrap || !tb) return;
    const n = Math.max(1, tb.querySelectorAll('th.col').length);
    tb.classList.toggle('hz', (wrap.clientWidth - 190) / n >= 62);
  }
  if (window.__rawFit) window.removeEventListener('resize', window.__rawFit);
  window.__rawFit = fit; window.addEventListener('resize', fit);
  draw();
  $('input[type=search]', el).addEventListener('input', e => { s.q = e.target.value; draw(); });
  $('#hz', el).addEventListener('change', e => { s.hide = e.target.checked; draw(); });
}

function matSourcingView(el) {
  const rows = D.matSourcing.rows;
  el.innerHTML = head('Material Recipe Sourcing', D.matSourcing.title) + '<div id="t"></div>';
  makeTable($('#t', el), {
    rows, noun: 'rows', placeholder: 'Search recipe, monster or location…',
    search: r => r.slice(0, 6).join(' '),
    group: r => r[0], groupCols: ['k'],
    stripe: r => { const m = matByName[String(r[0]).replace(/^Recipe: /, '')]; return m ? tierColor(m.tier) : ''; },
    filters: [
      { id: 'c', label: 'Chronicle', options: uniq(rows.map(r => r[1])).sort(byOrder(CHRON)), test: (r, v) => r[1] === v },
      { id: 's', label: 'Source type', options: uniq(rows.map(r => r[2])).sort(), test: (r, v) => r[2] === v },
    ],
    toggles: [{ id: 'notes', label: 'Show notes', def: false }],
    cols: [
      { key: 'k', label: 'Recipe scroll', get: r => r[0], html: r => link(r[0]) },
      { key: 'c', label: 'Chronicle', get: r => CHRON.indexOf(r[1]), html: r => `<span class="chr">${esc(r[1])}</span>` },
      { key: 's', label: 'Source', get: r => r[2], html: r => esc(r[2]) },
      { key: 'm', label: 'Monster / quest', get: r => r[3], cls: 'wrap', html: r => esc(r[3]) },
      { key: 'l', label: 'Location', get: r => r[4], cls: 'wrap', html: r => esc(r[4]) },
      { key: 'r', label: 'Rate', cls: 'num', get: r => r[5], sort: r => rateSort(r[5]), html: r => rate(r[5]) },
      { key: 'n', label: 'Notes', toggle: 'notes', nosort: true, cls: 'note', get: r => '', html: r => esc(r[6]) },
    ],
  }, st('matsrc'));
}

function equipmentView(el) {
  const items = D.equipment;
  el.innerHTML = head('Equipment Recipes', 'Every weapon, armor and jewelry recipe (D to S84). From B-grade up the recipe scroll is consumed on every craft and is not listed among the ingredients.') + '<div id="t"></div>';
  const cat = it => it.type.split(' — ')[0];
  makeTable($('#t', el), {
    rows: items, noun: 'items', pageSize: 100, placeholder: 'Search item, type or ingredient…',
    search: it => it.name + ' ' + it.type + ' ' + it.grade + ' ' + it.intro + ' ' + it.ing.map(i => i[0]).join(' '),
    stripe: it => gradeColor(it.grade),
    filters: [
      { id: 'g', label: 'Grade', options: gradeOpts, test: (it, v) => it.grade === v },
      { id: 'cat', label: 'Category', options: uniq(items.map(cat)), test: (it, v) => cat(it) === v },
      { id: 'c', label: 'Introduced', options: uniq(items.map(i => i.intro)).sort(byOrder(CHRON)), test: (it, v) => it.intro === v },
    ],
    cols: [
      { key: 'n', label: 'Item', get: it => it.name, html: it => link(it.name) },
      { key: 't', label: 'Type', get: it => it.type, cls: 'wrap hm', html: it => esc(it.type) },
      { key: 'g', label: 'Grade', get: it => GRADES.indexOf(it.grade), html: it => chip(it.grade, 'g-' + it.grade) },
      { key: 'i', label: 'Introduced', cls: 'hm', get: it => CHRON.indexOf(it.intro), html: it => `<span class="chr">${esc(it.intro)}</span>` },
      { key: 'sk', label: 'Skill lvl', cls: 'num hm', get: it => it.skill, html: it => num(it.skill) },
      { key: 'su', label: 'Success', cls: 'num hm', get: it => it.succ, html: it => pct(it.succ) },
      { key: 'mp', label: 'MP', cls: 'num hm', get: it => it.mp, html: it => num(it.mp) },
      { key: 'ing', label: 'Ingredients', nosort: true, get: () => '', html: it => ingList(it.ing) },
    ],
  }, st('equip'));
}

function equipSourcingView(el) {
  const rows = D.equipSourcing.rows;
  el.innerHTML = head('Equipment Sourcing', D.equipSourcing.title) + '<div id="t"></div>';
  makeTable($('#t', el), {
    rows, noun: 'rows', pageSize: 150, placeholder: 'Search item, monster or location…',
    search: r => r.slice(0, 7).join(' '),
    group: r => r[0], groupCols: ['k', 'g', 't'],
    stripe: r => (r[2] === 'Part' && (r[1] === 'S80' || r[1] === 'S84') ? 'var(--t-part)' : gradeColor(r[1])),
    filters: [
      { id: 'g', label: 'Grade', options: gradeOpts, test: (r, v) => r[1] === v },
      { id: 't', label: 'Type', options: ['Recipe', 'Part'], test: (r, v) => r[2] === v },
      { id: 'c', label: 'Chronicle', options: uniq(rows.map(r => r[3])).sort(byOrder(CHRON)), test: (r, v) => r[3] === v },
      { id: 's', label: 'Source type', options: uniq(rows.map(r => r[4])).sort(), test: (r, v) => r[4] === v },
    ],
    toggles: [{ id: 'notes', label: 'Show notes', def: false }],
    cols: [
      { key: 'k', label: 'Item / key material', get: r => r[0], cls: 'wrap', html: r => link(r[0]) },
      { key: 'g', label: 'Grade', get: r => GRADES.indexOf(r[1]), html: r => chip(r[1], 'g-' + r[1]) },
      { key: 't', label: 'Type', cls: 'hm', get: r => r[2], html: r => esc(r[2]) },
      { key: 'c', label: 'Chronicle', get: r => CHRON.indexOf(r[3]), html: r => `<span class="chr">${esc(r[3])}</span>` },
      { key: 's', label: 'Source', get: r => r[4], html: r => esc(r[4]) },
      { key: 'm', label: 'Monster / quest', get: r => r[5], cls: 'wrap', html: r => esc(r[5]) },
      { key: 'l', label: 'Location', get: r => r[6], cls: 'wrap', html: r => esc(r[6]) },
      { key: 'r', label: 'Rate', cls: 'num', get: r => r[7], sort: r => rateSort(r[7]), html: r => rate(r[7]) },
      { key: 'n', label: 'Notes', toggle: 'notes', nosort: true, cls: 'note', get: () => '', html: r => esc(r[8]) },
    ],
  }, st('equipsrc'));
}

function auditView(el) {
  const rows = D.audit;
  el.innerHTML = head('Audit Log', 'The original per-row check notes from each chronicle sourcing audit (the Notes column on the sourcing tabs holds only the current summary).') + '<div id="t"></div>';
  makeTable($('#t', el), {
    rows, noun: 'entries', pageSize: 150, placeholder: 'Search item, monster or note…',
    search: r => r.join(' '),
    group: r => r[0], groupCols: ['k', 'g'],
    stripe: r => gradeColor(r[1]),
    filters: [
      { id: 'g', label: 'Grade', options: gradeOpts, test: (r, v) => r[1] === v },
      { id: 'c', label: 'Chronicle', options: uniq(rows.map(r => r[2])).sort(byOrder(CHRON)), test: (r, v) => r[2] === v },
      { id: 's', label: 'Source type', options: uniq(rows.map(r => r[3])).sort(), test: (r, v) => r[3] === v },
    ],
    cols: [
      { key: 'k', label: 'Item / key material', get: r => r[0], cls: 'wrap', html: r => link(r[0]) },
      { key: 'g', label: 'Grade', get: r => GRADES.indexOf(r[1]), html: r => chip(r[1], 'g-' + r[1]) },
      { key: 'c', label: 'Chronicle', get: r => CHRON.indexOf(r[2]), html: r => `<span class="chr">${esc(r[2])}</span>` },
      { key: 's', label: 'Source', get: r => r[3], html: r => esc(r[3]) },
      { key: 'm', label: 'Monster / quest', get: r => r[4], cls: 'wrap', html: r => esc(r[4]) },
      { key: 'r', label: 'Rate', cls: 'num', get: r => r[5], html: r => esc(r[5]) },
      { key: 'n', label: 'Original check note', nosort: true, cls: 'note', get: () => '', html: r => esc(r[6]) },
    ],
  }, st('audit'));
}

function notesView(el) {
  const L = D.notes, isHead = (l, i) => {
    if (!l.t) return false;
    if (l.b) return true;
    const first = l.t.trim().split(/\s+/)[0];
    const prevBlank = i === 0 || !L[i - 1].t;
    return (prevBlank || /^(COLOR|SHEETS|GENERAL)$/.test(first)) && l.t.length <= 110 && /^[A-Z0-9][A-Z0-9\-]+$/.test(first) && first.length >= 2 && !/[.:]$/.test(first);
  };
  let html = '', buf = [];
  const flush = () => { if (buf.length) { html += `<div class="pre">${esc(buf.join('\n').replace(/^\n+|\n+$/g, ''))}</div>`; buf = []; } };
  L.slice(1).forEach((l, j) => {
    const i = j + 1;
    if (isHead(l, i)) { flush(); html += `<h3>${esc(l.t.trim())}</h3>`; } else buf.push(l.t);
  });
  flush();
  el.innerHTML = head(L[0].t.replace(/^Notes — /, 'Notes — ').trim(), '') + `<div class="card notes">${html}</div>`;
}

function soonView(title, lead, points) {
  return el => {
    el.innerHTML = head(title, lead) + `<div class="card"><p><b>Not connected yet.</b> This tab is part of the live, shared part of the site.</p><ul>${points.map(p => `<li>${esc(p)}</li>`).join('')}</ul></div>`;
  };
}


function testBanner() {
  return B.mode === 'local' ? '<div class="banner"><b>Test mode.</b> This copy is not connected to the shared database, so the warehouse below is saved in this browser only.</div>' : '';
}
function openModal(title, bodyHtml, okLabel, onOk, onCancel) {
  const m = $('#modal');
  m.innerHTML = `<div class="modalbox" role="dialog" aria-modal="true"><h2>${esc(title)}</h2>${bodyHtml}<div class="modalbtns"><button class="btn" data-mc>Cancel</button><button class="btn primary" data-mo>${esc(okLabel)}</button></div></div>`;
  m.hidden = false; m._cancel = onCancel || null;
  m.onclick = e => { if (e.target === m || e.target.closest('[data-mc]')) closeModal(true); else if (e.target.closest('[data-mo]')) { closeModal(); onOk(); } };
  $('[data-mo]', m).focus();
}
function closeModal(cancelled) { const m = $('#modal'), c = m._cancel; m.hidden = true; m.innerHTML = ''; m._cancel = null; if (cancelled && c) c(); }
let toastTimer = null;
function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 3500); }
function rerender() { const el = $('#view'); el.innerHTML = ''; viewById[current].render(el); }

/* ------------------------------------------------- loading, login, live refresh */
const draft = {};   // typed-but-unsaved warehouse numbers: item -> { val, base }  (base = the stock the person saw)
async function reload() { try { store = await B.load(); loaded = true; loadErr = null; } catch (e) { loadErr = e; } }
function ensureLoaded() { if (!loading) loading = reload().then(() => { loading = null; }); return loading; }
let refreshTimer = null;
function scheduleRefresh() {
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(async () => { if (B.mode === 'shared' && !session) return; await reload(); if (onData && !loadErr) onData(); }, 350);
}
function errText(e) {
  if (e && e.notAllowed) return 'This account is not on the allowed list. Ask the owner to add your e-mail.';
  return (e && e.message) || String(e);
}
function gated(fn) {
  return el => {
    if (B.mode === 'shared' && !session) { loginView(el); return; }
    const cv = current;
    const draw = () => {
      if (current !== cv || $('#view') !== el) return;
      el.innerHTML = '';
      if (loadErr) errorView(el); else fn(el);
    };
    if (!loaded) { el.innerHTML = '<div class="card"><p class="lead" style="margin:0">Loading the warehouse…</p></div>'; ensureLoaded().then(draw); } else fn(el);
  };
}
function errorView(el) {
  el.innerHTML = `<div class="card"><h2 style="margin-top:0">The warehouse could not be loaded</h2><p class="lead">${esc(errText(loadErr))}</p><button class="btn" id="retry">Try again</button></div>`;
  $('#retry', el).addEventListener('click', () => { loadErr = null; loaded = false; rerender(); });
}
function loginView(el) {
  el.innerHTML = head('Sign in', 'The clan tools (Warehouse, Requirements Check, History) are shared by two people. Sign in with the account you were given.') + `
    <form class="card login" id="lf">
      <label>E-mail<input type="email" id="le" autocomplete="username" required></label>
      <label>Password<input type="password" id="lp" autocomplete="current-password" required></label>
      <div class="commitbar"><button class="btn primary" type="submit">Sign in</button><span class="err" id="lerr" role="alert"></span></div>
    </form>`;
  $('#lf', el).addEventListener('submit', async e => {
    e.preventDefault();
    const btn = $('button[type=submit]', el), err = $('#lerr', el);
    btn.disabled = true; err.textContent = '';
    let r; try { r = await B.signIn($('#le', el).value.trim(), $('#lp', el).value); } catch (x) { r = { error: errText(x) }; }
    if (r.error) { btn.disabled = false; err.textContent = /invalid login/i.test(r.error) ? 'Wrong e-mail or password.' : r.error; }
  });
  $('#le', el).focus();
}
let unpres = null, onlineNow = [];
function renderOnline() {
  const el = $('#online');
  if (B.mode !== 'shared' || !session || !B.presence) { el.hidden = true; return; }
  const me = short(session.email), others = onlineNow.filter(n => n !== me);
  el.hidden = false; el.classList.toggle('alone', !others.length);
  el.title = others.length ? 'Online now: ' + [me].concat(others).map(nice).join(', ') : 'Only you are online right now';
  el.innerHTML = others.length ? `<i></i><span>${esc(others.map(nice).join(', '))} online</span>` : '<i></i><span>Only you online</span>';
}
function stopPresence() { if (unpres) { unpres(); unpres = null; } onlineNow = []; renderOnline(); }
function startPresence() { stopPresence(); if (B.mode === 'shared' && session && B.presence) unpres = B.presence(list => { onlineNow = list; renderOnline(); }); }
/* only accounts on the allowed list take part in the online status (the warehouse has to load first) */
function maybePresence() { if (B.mode !== 'shared' || !session) return; ensureLoaded().then(() => { if (session && loaded && !unpres) startPresence(); }); }
function renderWho() {
  renderOnline();
  const w = $('#who');
  if (B.mode === 'local') w.innerHTML = '<span class="tag">test mode</span>';
  else if (session) w.innerHTML = `<span class="whoname" title="${esc(session.email)}">${esc(nice(session.email))}</span><button class="iconbtn" id="signout">Sign out</button>`;
  else w.innerHTML = '<a class="iconbtn" href="#/home">Sign in</a>';
}
function setSession(user) {
  session = user; loaded = false; loadErr = null; loading = null; store = { stock: {}, learned: {}, meta: {} };
  Object.keys(draft).forEach(k => delete draft[k]);
  if (unsub) { unsub(); unsub = null; }
  if (user || B.mode === 'local') unsub = B.subscribe(scheduleRefresh);
  if (user) maybePresence(); else stopPresence();
  renderWho();
  if (viewById[current] && (viewById[current].group === 'Clan tools' || current === 'home')) rerender();
}
const conflictLines = list => list.map(c => `<div class="rowline"><span>${esc(c.item)}</span><span class="k">now <b>${num(c.current)}</b>${c.by ? ` · changed by ${esc(nice(c.by))} ${esc(when(c.at))}` : ''} <span class="hint">(you saw ${num(c.expected)}${c.new != null ? `, you wanted ${num(c.new)}` : ''})</span></span></div>`).join('');

/* ---------------------------------------------------------------- Warehouse */
function warehouseView(el) {
  const s = st('wh');
  const rows = D.catalogue.map(([cat, name, grade, intro]) => ({ cat, name, grade, intro }));
  const tier = cat => (cat === 'Raw Material' ? 'raw' : /^Tier (\d)/.test(cat) ? cat.slice(5) : cat === 'Recipe' ? 'recipe' : cat === 'Equipment Part' ? 'part' : 'crystal');
  const root = document.createElement('div'); root.className = 'whwrap' + (s.stk ? ' stocktake' : ''); el.appendChild(root);
  root.innerHTML = head('Warehouse', 'Type the quantity you actually hold in the game warehouse and press Update — the site stores that number, so there is no need to work out what to add. Both of you see the same numbers.') + testBanner() + `
    <div class="controls">
      <button class="btn" data-act="stk" aria-pressed="${!!s.stk}">Stock-take mode: ${s.stk ? 'on' : 'off'}</button>
      <button class="btn" data-act="csv">Download CSV</button>
      ${B.canDemo ? `<button class="btn" data-act="demo">Load demo stock</button><button class="btn" data-act="reset">Reset all to 0</button><button class="btn" data-act="export">Export JSON</button><button class="btn" data-act="import">Import JSON</button><input type="file" id="imp" accept="application/json,.json" hidden>` : ''}
    </div>
    <p class="hint" id="stkhint"${s.stk ? '' : ' hidden'}>Stock-take: type the counts from the game, press Enter to jump to the next row, then press <b>Save all</b>. Filters and search keep what you typed.</p>
    <div id="t"></div><div class="draftbar" id="draftbar" hidden></div>`;
  const lcHtml = n => { const m = store.meta[n]; return m && m.at ? `${esc(nice(m.by))} · ${esc(when(m.at))}` : ''; };
  const val = inp => Math.min(2000000000, Math.max(0, Math.floor(Number(inp.value)) || 0));
  const cell = r => {
    const d = draft[r.name], cur = store.stock[r.name] || 0, v = d ? d.val : cur, base = d ? d.base : cur, dl = v - cur;
    return `<span class="setto"><input type="number" min="0" step="1" inputmode="numeric" value="${v}" data-in="${esc(r.name)}" data-base="${base}" aria-label="New quantity of ${esc(r.name)}"><button class="btn" data-upd="${esc(r.name)}">Update</button><span class="delta ${dl > 0 ? 'up' : dl < 0 ? 'down' : ''}">${dl ? `(${sgn(dl)})` : ''}</span></span>`;
  };
  makeTable($('#t', root), {
    rows, noun: 'materials', pageSize: 100, placeholder: 'Search material…',
    search: r => r.name + ' ' + r.cat,
    stripe: r => tierColor(tier(r.cat)),
    filters: [
      { id: 'cat', label: 'Category', options: uniq(rows.map(r => r.cat)), test: (r, v) => r.cat === v },
      { id: 'g', label: 'Grade', options: uniq(rows.map(r => r.grade)).filter(g => g !== '—').sort(byOrder(GRADES)), test: (r, v) => r.grade === v },
      { id: 'i', label: 'Introduced', options: uniq(rows.map(r => r.intro)).filter(x => x !== '—').sort(byOrder(CHRON)), test: (r, v) => r.intro === v },
    ],
    cols: [
      { key: 'cat', label: 'Category', cls: 'hm', get: r => r.cat, html: r => esc(r.cat) },
      { key: 'n', label: 'Material', get: r => r.name, cls: 'wrap', html: r => link(r.name) },
      { key: 'g', label: 'Grade', cls: 'hm', get: r => GRADES.indexOf(r.grade), html: r => (r.grade === '—' ? '<span style="color:var(--muted)">—</span>' : chip(r.grade, 'g-' + r.grade)) },
      { key: 'i', label: 'Introduced', cls: 'hm', get: r => CHRON.indexOf(r.intro), html: r => (r.intro === '—' ? '<span style="color:var(--muted)">—</span>' : `<span class="chr">${esc(r.intro)}</span>`) },
      { key: 'q', label: 'In stock', cls: 'num', get: r => store.stock[r.name] || 0, html: r => `<span data-q="${esc(r.name)}">${num(store.stock[r.name] || 0)}</span>` },
      { key: 'lc', label: 'Last change', cls: 'hm', get: r => (store.meta[r.name] || {}).at || '', html: r => `<span class="when" data-lc="${esc(r.name)}">${lcHtml(r.name)}</span>` },
      { key: 'set', label: 'Set to', nosort: true, get: () => '', html: cell },
    ],
  }, s);

  const showDelta = inp => {
    const n = inp.dataset.in, d = draft[n] ? draft[n].val - (store.stock[n] || 0) : 0, out = inp.parentNode.querySelector('.delta');
    out.textContent = d ? `(${sgn(d)})` : ''; out.className = 'delta ' + (d > 0 ? 'up' : d < 0 ? 'down' : '');
  };
  const drawBar = () => {
    const bar = $('#draftbar', root), n = Object.keys(draft).length;
    bar.hidden = !n;
    bar.innerHTML = n ? `<span><b>${n}</b> unsaved change${n > 1 ? 's' : ''}</span><button class="btn primary" data-act="saveall">Save all</button><button class="btn" data-act="discard">Discard</button>` : '';
  };
  function refreshCells() {
    root.querySelectorAll('[data-q]').forEach(c => { c.textContent = num(store.stock[c.dataset.q] || 0); });
    root.querySelectorAll('[data-lc]').forEach(c => { c.innerHTML = lcHtml(c.dataset.lc); });
    root.querySelectorAll('[data-in]').forEach(inp => { const n = inp.dataset.in, cur = store.stock[n] || 0; if (!draft[n]) { inp.value = cur; inp.dataset.base = cur; } else inp.dataset.base = draft[n].base; showDelta(inp); });
    drawBar();
  }
  onData = refreshCells;

  async function saveBatch(changes, action, okMsg) {
    let r;
    try { r = await B.apply(changes, action); } catch (e) { toast('Could not save — ' + errText(e)); return false; }
    if (r.ok) { changes.forEach(c => delete draft[c.item]); await reload(); refreshCells(); toast(okMsg || `Saved ${changes.length} change${changes.length > 1 ? 's' : ''}.`); return true; }
    await reload(); refreshCells();
    openModal('Someone changed this first', `<p class="lead">The warehouse changed while you were typing:</p><div class="changes">${conflictLines(r.conflicts)}</div><p class="hint">Nothing was saved. Overwrite with your numbers anyway?</p>`, 'Set anyway', async () => {
      const now = Object.fromEntries(r.conflicts.map(c => [c.item, c.current]));
      await saveBatch(changes.map(c => (c.item in now ? Object.assign({}, c, { expected: now[c.item] }) : c)), action, okMsg);
    }, () => { r.conflicts.forEach(c => { if (draft[c.item]) draft[c.item].base = c.current; }); refreshCells(); });
    return false;
  }

  root.addEventListener('input', e => {
    const inp = e.target.closest('[data-in]'); if (!inp) return;
    const n = inp.dataset.in, cur = store.stock[n] || 0;
    if (inp.value === '') delete draft[n];
    else { const v = val(inp); if (v === cur) delete draft[n]; else (draft[n] = draft[n] || { base: Number(inp.dataset.base) }).val = v; }
    showDelta(inp); drawBar();
  });
  root.addEventListener('keydown', e => {
    const inp = e.target.closest('[data-in]'); if (!inp || e.key !== 'Enter') return;
    if (s.stk) { const all = [...root.querySelectorAll('[data-in]')], nx = all[all.indexOf(inp) + 1]; if (nx) { nx.focus(); nx.select(); } }
    else inp.parentNode.querySelector('[data-upd]').click();
  });
  root.addEventListener('click', async e => {
    const u = e.target.closest('[data-upd]');
    if (u) {
      const name = u.dataset.upd, inp = u.parentNode.querySelector('[data-in]');
      if (inp.value === '') { toast('Type a number first.'); return; }
      const v = val(inp), cur = store.stock[name] || 0, base = Number(inp.dataset.base);
      if (v === cur && base === cur) { delete draft[name]; inp.value = cur; showDelta(inp); drawBar(); toast('No change — that is already the stored number.'); return; }
      u.disabled = true;
      const ok = await saveBatch([{ item: name, new: v, expected: base }], 'set', `${name}: ${num(cur)} → ${num(v)}`);
      u.disabled = false;
      if (ok) { u.textContent = 'Saved ✓'; setTimeout(() => (u.textContent = 'Update'), 900); }
      return;
    }
    const act = e.target.closest('[data-act]'); if (!act) return;
    const a = act.dataset.act;
    if (a === 'stk') {
      s.stk = !s.stk; root.classList.toggle('stocktake', s.stk); act.textContent = 'Stock-take mode: ' + (s.stk ? 'on' : 'off'); act.setAttribute('aria-pressed', String(!!s.stk)); $('#stkhint', root).hidden = !s.stk;
    } else if (a === 'saveall') {
      const changes = Object.entries(draft).map(([item, d]) => ({ item, new: d.val, expected: d.base }));
      if (!changes.length) return;
      act.disabled = true; await saveBatch(changes, 'stocktake', `Stock-take saved (${changes.length} item${changes.length > 1 ? 's' : ''}).`); act.disabled = false;
    } else if (a === 'discard') {
      Object.keys(draft).forEach(k => delete draft[k]); refreshCells();
    } else if (a === 'csv') {
      const q = v => (/[",\r\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v));
      const lines = [['Category', 'Material', 'Grade', 'Introduced', 'In stock']].concat(rows.map(r => [r.cat, r.name, r.grade === '—' ? '' : r.grade, r.intro === '—' ? '' : r.intro, store.stock[r.name] || 0]));
      download('l2-warehouse-' + new Date().toISOString().slice(0, 10) + '.csv', '﻿' + lines.map(l => l.map(q).join(',')).join('\r\n'), 'text/csv;charset=utf-8');
    } else if (a === 'demo') {
      const st0 = {}; rows.forEach(r => { const t = tier(r.cat), v = t === 'raw' ? 3000 : t === 'crystal' ? 5000 : t === 'part' ? 30 : t === 'recipe' ? 3 : 0; if (v) st0[r.name] = v; });
      await B.replaceAll(st0, store.learned); await reload(); rerender(); toast('Demo stock loaded.');
    } else if (a === 'reset') {
      if (confirm('Set every quantity back to 0 and forget which recipes are learned?')) { await B.replaceAll({}, {}); Object.keys(draft).forEach(k => delete draft[k]); await reload(); rerender(); toast('Warehouse cleared.'); }
    } else if (a === 'export') {
      download('l2-test-warehouse.json', JSON.stringify(await B.dump(), null, 1), 'application/json');
    } else if (a === 'import') $('#imp', root).click();
  });
  const imp = $('#imp', root);
  if (imp) imp.addEventListener('change', e => {
    const f = e.target.files[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = async () => {
      try {
        const j = JSON.parse(rd.result);
        if (!j || typeof j.stock !== 'object') throw new Error('not a warehouse file');
        await B.replaceAll(j.stock, j.learned || {}); await reload(); rerender(); toast('Warehouse imported.');
      } catch (err) { toast('Could not read that file.'); }
    };
    rd.readAsText(f);
  });
}
function download(name, text, type) {
  const blob = new Blob([text], { type }), url = URL.createObjectURL(blob), l = document.createElement('a');
  l.href = url; l.download = name; document.body.appendChild(l); l.click(); l.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ------------------------------------------------------------------ History */
const ACTION = { set: 'Update', stocktake: 'Stock-take', commit: 'Commit', learn: 'Recipe learned', unlearn: 'Recipe un-learned' };
function historyView(el) {
  const s = st('hist');
  el.innerHTML = head('History', 'Every change to the shared warehouse — who, what and when. Newest first (last 1,000 entries). A Commit shows one line per material it used.') + testBanner() + '<div id="t"><p class="lead">Loading…</p></div>';
  const draw = list => {
    const host = $('#t', el); if (!host || current !== 'history') return;
    const rows = list.map(h => ({ id: h.id, at: h.at, by: nice(h.by), action: ACTION[h.action] || h.action, item: h.item, o: h.old_qty, n: h.new_qty, label: h.label || '' }));
    makeTable(host, {
      rows, noun: 'entries', pageSize: 150, placeholder: 'Search item, person or note…',
      search: r => [r.item, r.by, r.action, r.label].join(' '),
      filters: [
        { id: 'p', label: 'Person', options: uniq(rows.map(r => r.by)).sort(), test: (r, v) => r.by === v },
        { id: 'a', label: 'Action', options: uniq(rows.map(r => r.action)), test: (r, v) => r.action === v },
      ],
      cols: [
        { key: 'at', label: 'When', get: r => r.id, html: r => esc(when(r.at)) },
        { key: 'by', label: 'Who', get: r => r.by, html: r => esc(r.by) },
        { key: 'a', label: 'Action', get: r => r.action, html: r => esc(r.action) },
        { key: 'i', label: 'Item', cls: 'wrap', get: r => r.item, html: r => link(r.item) },
        { key: 'c', label: 'Change', cls: 'num', get: r => (r.n == null ? 0 : r.n - r.o), html: r => (r.n == null ? '<span style="color:var(--muted)">—</span>' : `${num(r.o)} → <b>${num(r.n)}</b> <span class="${r.n < r.o ? 'neg' : 'pos'}">(${sgn(r.n - r.o)})</span>`) },
        { key: 'l', label: 'Note', cls: 'hm wrap', get: r => r.label, html: r => esc(r.label) },
      ],
    }, s);
  };
  const fetchAll = () => B.history(1000).then(draw).catch(e => { const h = $('#t', el); if (h) h.innerHTML = `<p class="lead">${esc(errText(e))}</p>`; });
  onData = fetchAll; fetchAll();
}

function treeHtml(n) {
  const mat = E.matByName[n.name];
  let tags = '';
  if (n.fromStock > 0) tags += `<span class="tag">in stock ${num(n.fromStock)}</span>`;
  if (mat && n.crafts > 0) tags += `<span class="tag craft">craft ${num(n.crafted)}${n.makes > n.crafted ? ` · ${n.crafts} batch${n.crafts > 1 ? 'es' : ''} make ${num(n.makes)}` : n.crafts !== n.crafted ? ` · ${n.crafts} crafts` : ''}</span>`;
  if (!mat && n.missing > 0) tags += `<span class="tag bad">need ${num(n.missing)} more</span>`;
  const kids = n.children.length ? `<ul>${n.children.map(treeHtml).join('')}</ul>` : '';
  return `<li>${link(n.name)} <span class="x">×${num(n.needed)}</span> ${tags}${kids}</li>`;
}

function resultHtml(R) {
  const it = R.item, bad = R.rows.filter(r => !r.ok).length;
  const rowsHtml = R.rows.map(r => `<tr>
      <td>${link(r.name)}${r.isScroll ? ' <span class="chip t-recipe">scroll</span>' : ''}</td>
      <td class="num">${num(r.needed)}</td><td class="num">${num(r.warehouse)}</td>
      <td class="num">${r.short ? num(r.short) : '—'}</td>
      <td class="num">${r.short && r.kind === 'Material' ? num(r.craftable) : '—'}</td>
      <td class="num bal ${r.balance >= 0 ? 'ok' : 'bad'}">${sgn(r.balance)}</td>
      <td><span class="st ${r.ok ? 'ok' : 'bad'}">${r.ok ? 'OK' : 'SHORT'}</span></td></tr>`).join('');
  const crafted = R.rows.filter(r => r.node.crafts > 0);
  const tree = crafted.length
    ? `<ul class="tree">${crafted.map(r => treeHtml(r.node)).join('')}</ul>`
    : `<p class="lead">${R.ok ? 'Everything is already in the warehouse — nothing to craft.' : 'None of the missing ingredients can be crafted. They have to be farmed (see the list below).'}</p>`;
  const order = ['Raw material', 'Equipment part', 'Recipe scroll', 'Crystal / Gemstone'];
  const shop = R.shopping.slice().sort((a, b) => (order.indexOf(a.kind) + 1 || 9) - (order.indexOf(b.kind) + 1 || 9) || a.name.localeCompare(b.name));
  const scrollNote = R.scroll
    ? (R.scroll.consumed
      ? `Recipe scroll: consumed on every craft${R.scroll.learned ? '' : ', plus 1 more to learn the recipe the first time'}.`
      : `Recipe scroll: reusable at ${it.grade}-grade — 1 scroll to learn it, then never again.`)
    : (E.REUSABLE.has(it.grade) ? 'Recipe already learned (reusable at this grade).' : '');
  return `<div class="reqgrid">
    <div class="card">
      <div class="reqhead"><div><h2>${esc(it.name)} ×${R.qty}</h2>
        <div class="meta">${chip(it.grade, 'g-' + it.grade)} <span class="chr">${esc(it.intro)}</span> ${esc(it.type)} · Skill lvl ${num(it.skill)} · Success ${pct(it.succ)} · MP ${num(it.mp)}</div></div>
        <div class="verdict ${R.ok ? 'ok' : 'bad'}">${R.ok ? 'Ready to craft' : bad + ' ingredient' + (bad > 1 ? 's' : '') + ' short'}</div></div>
      <div class="prog"><span>${R.rows.length - bad} of ${R.rows.length} ready</span><div class="track"><div class="fill${R.ok ? ' done' : ''}" style="width:${R.rows.length ? Math.round((R.rows.length - bad) / R.rows.length * 100) : 100}%"></div></div><span>${R.rows.length ? Math.round((R.rows.length - bad) / R.rows.length * 100) : 100}%</span></div>
      <div class="tablewrap flat"><table class="req"><thead><tr><th>Ingredient</th><th class="num">Needed</th><th class="num">Warehouse</th><th class="num">To craft</th><th class="num" title="How many of the missing units you could make right now from materials in stock">Craftable</th><th class="num">Balance</th><th>Status</th></tr></thead><tbody>${rowsHtml}</tbody></table></div>
      ${scrollNote ? `<p class="hint">${esc(scrollNote)}</p>` : ''}
      <div class="commitbar"><button class="btn primary" id="commit"${R.ok ? '' : ' disabled'}>Commit</button>
        <span class="hint">${R.ok ? 'Subtracts these materials from the warehouse (you will see the changes first).' : 'Commit unlocks when every balance is green.'}</span></div>
    </div>
    <div class="card"><h3 style="margin-top:0">Sub-materials to craft</h3>${tree}</div>
  </div>
  ${shop.length ? `<div class="card"><h3 style="margin-top:0">Still to get (${shop.length})</h3>
    <div class="tablewrap flat"><table><thead><tr><th>Item</th><th>Kind</th><th class="num">Required</th><th class="num">In warehouse</th><th class="num">Still needed</th></tr></thead><tbody>
    ${shop.map(l => `<tr><td>${link(l.name)}</td><td>${chip(l.kind, kindCls(l.kind))}</td><td class="num">${num(l.required)}</td><td class="num">${num(l.warehouse)}</td><td class="num bal bad">${num(l.missing)}</td></tr>`).join('')}
    </tbody></table></div></div>` : ''}`;
}

function reqView(el) {
  const s = st('req');
  s.grade = s.grade || gradeOpts[0]; s.qty = s.qty || 1; s.filter = s.filter || '';
  el.innerHTML = head('Requirements Check', 'Pick a grade, an item and how many to craft. Each ingredient is checked against the warehouse; anything short is broken down into the sub-materials you still need.') + testBanner() + `
    <div class="card reqctl">
      <label>Grade<select id="rg">${gradeOpts.map(g => `<option${g === s.grade ? ' selected' : ''}>${g}</option>`).join('')}</select></label>
      <label class="grow">Item<span class="pair"><input type="search" id="rf" placeholder="Filter items…" value="${esc(s.filter)}"><select id="ri"></select></span></label>
      <label>Quantity<input type="number" id="rq" min="1" max="999" step="1" value="${s.qty}"></label>
      <label class="tog"><input type="checkbox" id="rl"> Recipe already learned</label>
    </div><div id="rout"></div>`;
  const ri = $('#ri', el), out = $('#rout', el);
  function fillItems() {
    const q = lc(s.filter), items = D.equipment.filter(i => i.grade === s.grade && (!q || lc(i.name + ' ' + i.type).includes(q)));
    const groups = {}; items.forEach(i => (groups[i.type.split(' — ')[0]] = groups[i.type.split(' — ')[0]] || []).push(i));
    ri.innerHTML = items.length ? Object.entries(groups).map(([g, l]) => `<optgroup label="${esc(g)}">${l.map(i => `<option value="${esc(i.name)}">${esc(i.name)}</option>`).join('')}</optgroup>`).join('') : '<option value="">No match</option>';
    if (!items.some(i => i.name === s.item)) s.item = items.length ? items[0].name : null;
    ri.value = s.item || '';
  }
  function render() {
    const it = s.item && E.itemByName[s.item];
    if (!it) { out.innerHTML = '<p class="lead">No item matches the filter.</p>'; return; }
    const learned = !!store.learned[it.name]; $('#rl', el).checked = learned;
    out.innerHTML = resultHtml(E.resolve(it.name, s.qty, store.stock, learned));
  }
  fillItems(); render();
  $('#rg', el).addEventListener('change', e => { s.grade = e.target.value; fillItems(); render(); });
  $('#rf', el).addEventListener('input', e => { s.filter = e.target.value; fillItems(); render(); });
  ri.addEventListener('change', e => { s.item = e.target.value; render(); });
  $('#rq', el).addEventListener('input', e => { const v = Math.floor(Number(e.target.value)); if (v >= 1) { s.qty = Math.min(v, 999); render(); } });
  $('#rl', el).addEventListener('change', async e => {
    const chk = e.target.checked, prev = !!store.learned[s.item], item = s.item;
    if (chk) store.learned[item] = true; else delete store.learned[item];
    render();
    try { await B.setLearned(item, chk); } catch (err) { if (prev) store.learned[item] = true; else delete store.learned[item]; render(); toast('Could not save — ' + errText(err)); }
  });
  onData = render;
  out.addEventListener('click', e => {
    if (!e.target.closest('#commit')) return;
    const R = E.resolve(s.item, s.qty, store.stock, !!store.learned[s.item]);
    if (!R.ok) return;
    const list = R.changes.slice().sort((a, b) => a.name.localeCompare(b.name)).map(c => `<div class="rowline"><span>${esc(c.name)}</span><span class="k">${num(c.before)} → <b>${num(c.after)}</b> <span class="${c.after < c.before ? 'neg' : 'pos'}">(${sgn(c.after - c.before)})</span></span></div>`).join('');
    openModal(`Commit ${R.item.name} ×${R.qty}?`,
      `<p class="lead">These warehouse quantities will change:</p><div class="changes">${list || '<p class="lead">No stock changes.</p>'}</div>${store.learned[s.item] ? '' : '<p class="hint">This recipe will be marked as learned.</p>'}`,
      'Commit', async () => {
        const btn = $('#commit', el); if (btn) btn.disabled = true;
        const changes = R.changes.map(c => ({ item: c.name, new: c.after, expected: c.before }));
        let res;
        try { res = await B.apply(changes, 'commit', `Commit ${R.item.name} ×${R.qty}`, R.item.name); } catch (err) { toast('Could not commit — ' + errText(err)); render(); return; }
        await reload();
        if (current === 'reqcheck') render();
        if (!res.ok) {
          openModal('The warehouse changed — nothing was committed', `<p class="lead">Someone changed these numbers after you ran the check:</p><div class="changes">${conflictLines(res.conflicts)}</div><p class="hint">The check below has been refreshed with the new numbers. Look it over and commit again if it is still green.</p>`, 'OK', () => {});
          return;
        }
        toast(`Committed ${R.item.name} ×${R.qty}.`);
      });
  });
}

/* ---------------------------------------------------------------------- Home */
function homeView(el) {
  const quick = `<div class="cols2 quick">
    <a class="card tl" href="#/warehouse"><h3>Warehouse</h3><p>What the clan holds right now. Type what the game shows; both of you see it.</p></a>
    <a class="card tl" href="#/reqcheck"><h3>Requirements Check</h3><p>Pick an item and see what is missing, what to craft first, then Commit.</p></a>
    <a class="card tl" href="#/equip"><h3>Equipment Recipes</h3><p>Every recipe and what it is made from.</p></a>
    <a class="card tl" href="#/start"><h3>Start Here</h3><p>How the workbook and the colours work.</p></a></div>`;
  if (B.mode === 'shared' && !session) { loginView(el); el.insertAdjacentHTML('beforeend', quick); return; }   // sign in right here, so you stay on Home afterwards
  gated(inner)(el);
  function inner(host) {
    const who = session ? nice(session.email) : '';
    host.innerHTML = head(who ? 'Welcome, ' + who : 'Welcome', 'Where the clan stands today: what is in the warehouse, what you can craft right now and what changed last.') + '<div id="hm"></div>' + quick;
    const box = $('#hm', host); let hist = [];
    const draw = () => {
      if (!$('#hm', host)) return;
      const stockItems = Object.entries(store.stock).filter(([, q]) => q > 0), units = stockItems.reduce((a, [, q]) => a + q, 0);
      const lEq = D.equipment.filter(i => store.learned[i.name]), lMat = D.materials.filter(m => store.learned[m.name]);
      const res = lEq.map(i => ({ it: i, R: E.resolve(i.name, 1, store.stock, true) }));
      const ready = res.filter(x => x.R.ok), near = res.filter(x => !x.R.ok).map(x => ({ it: x.it, n: x.R.rows.filter(r => !r.ok).length })).sort((a, b) => a.n - b.n || a.it.name.localeCompare(b.it.name)).slice(0, 5);
      const last = hist[0];
      const req = (it, extra) => `<div class="rowline"><button class="lnk" data-req="${esc(it.name)}">${esc(it.name)}</button> ${chip(it.grade, 'g-' + it.grade)}<span class="k">${extra}</span></div>`;
      const acts = hist.slice(0, 8).map(h => `<div class="rowline"><span class="when">${esc(when(h.at))}</span> <b>${esc(nice(h.by))}</b> ${esc(ACTION[h.action] || h.action)} ${link(h.item)}<span class="k">${typeof h.old_qty === 'number' ? `${num(h.old_qty)} → <b>${num(h.new_qty)}</b>` : ''}</span></div>`).join('');
      box.innerHTML = `<div class="tiles">
        <div class="card tile"><b>${num(stockItems.length)}</b><span>materials in the warehouse</span><small>${num(units)} units in total</small></div>
        <div class="card tile"><b>${num(lEq.length)}</b><span>equipment recipes learned</span><small>${num(lMat.length)} material recipes ticked</small></div>
        <div class="card tile"><b>${num(ready.length)}</b><span>ready to craft now</span><small>from the recipes you have learned</small></div>
        <div class="card tile"><b class="sm">${last ? esc(nice(last.by)) : '—'}</b><span>last change</span><small>${last ? esc(when(last.at)) + ' · ' + esc(ACTION[last.action] || last.action) : 'nothing yet'}</small></div></div>
        <div class="cols2">
          <div class="card"><h3 style="margin-top:0">Ready to craft now</h3>${ready.length ? ready.slice(0, 12).map(x => req(x.it, 'ready')).join('') + (ready.length > 12 ? `<p class="hint">…and ${ready.length - 12} more.</p>` : '') : '<p class="lead" style="margin:0">Nothing yet. Fill the warehouse and tick “Recipe already learned” in Requirements Check; learned recipes show up here when you hold everything for them.</p>'}
            ${near.length ? `<h4 style="margin:16px 0 6px;font:700 12px var(--sans);text-transform:uppercase;letter-spacing:.1em;color:var(--muted)">Closest to ready</h4>${near.map(x => req(x.it, x.n + ' short')).join('')}` : ''}</div>
          <div class="card"><h3 style="margin-top:0">Latest changes</h3>${acts || '<p class="lead" style="margin:0">No changes yet.</p>'}
            <p class="hint"><a href="#/history">Open the full History →</a></p></div></div>`;
    };
    const refresh = () => B.history(8).then(h => { hist = h; draw(); }, () => draw());
    host.onclick = e => {
      const b = e.target.closest('[data-req]'); if (!b) return;
      const it = itemByName[b.dataset.req], s = st('req'); s.grade = it.grade; s.item = it.name; s.filter = ''; go('reqcheck');
    };
    onData = refresh; draw(); refresh();
  }
}

const VIEWS = [
  { id: 'home', label: 'Home', group: '', render: homeView },
  { id: 'warehouse', label: 'Warehouse', group: 'Clan tools', render: gated(warehouseView) },
  { id: 'reqcheck', label: 'Requirements Check', group: 'Clan tools', render: gated(reqView) },
  { id: 'history', label: 'History', group: 'Clan tools', render: gated(historyView) },
  { id: 'start', label: 'Start Here', group: 'Reference', render: startView },
  { id: 'materials', label: 'Material Recipes', group: 'Reference', render: materialsView },
  { id: 'raw', label: 'Raw Material Breakdown', group: 'Reference', render: rawView },
  { id: 'matsrc', label: 'Material Recipe Sourcing', group: 'Reference', render: matSourcingView },
  { id: 'equip', label: 'Equipment Recipes', group: 'Reference', render: equipmentView },
  { id: 'equipsrc', label: 'Equipment Sourcing', group: 'Reference', render: equipSourcingView },
  { id: 'notes', label: 'Notes', group: 'Reference', render: notesView, hidden: true },
  { id: 'audit', label: 'Audit Log', group: 'Reference', render: auditView, hidden: true },
];
const viewById = Object.fromEntries(VIEWS.map(v => [v.id, v]));

/* ------------------------------------------------------------------- drawer */
function sourcesHtml(rows, kind) {
  if (!rows || !rows.length) return '<p class="lead">No sourcing rows in the workbook.</p>';
  return rows.map(r => {
    const [chr, type, mon, loc, rt, note] = kind === 'eq' ? [r[3], r[4], r[5], r[6], r[7], r[8]] : [r[1], r[2], r[3], r[4], r[5], r[6]];
    return `<div class="src"><span class="rate">${rate(rt)}</span><span class="chr">${esc(chr)}</span> <b>${esc(type)}</b> · ${esc(mon)}<br><span style="color:var(--muted)">${esc(loc)}</span>${note ? `<div class="note">${esc(note)}</div>` : ''}</div>`;
  }).join('');
}
function pills(list) { return `<div class="pills">${list.map(n => link(n)).join('')}</div>`; }
function usedBlock(name) {
  const m = usedInMats[name] || [], it = usedInItems[name] || [];
  let h = '';
  if (m.length) h += `<h4>Used in materials (${m.length})</h4>${pills(m.map(x => x.name))}`;
  if (it.length) h += `<h4>Used in equipment (${it.length})</h4>${pills(it.slice(0, 60).map(x => x.name))}${it.length > 60 ? `<p class="lead">…and ${it.length - 60} more — search the Equipment Recipes tab for “${esc(name)}”.</p>` : ''}`;
  return h;
}
function detail(name) {
  const k = kindOf(name), it = itemByName[name], mat = matByName[name];
  let h = `<h2>${esc(name)}</h2><div class="meta">${chip(k.label, k.cls)}`;
  let body = '';
  if (it) {
    h += ` <span class="chr">${esc(it.intro)}</span></div><div class="meta">${esc(it.type)}</div><div class="meta">Skill lvl ${num(it.skill)} · Success ${pct(it.succ)} · MP ${num(it.mp)}</div>`;
    body += `<h4>Ingredients</h4>${it.ing.map(([n, q]) => `<div class="rowline"><span class="q">${num(q)}×</span>${link(n, isCurrency(n) ? 'dim' : '')}<span class="k">${chip(kindOf(n).label, kindOf(n).cls)}</span></div>`).join('')}`;
    body += `<h4>Recipe scroll — where to get it</h4>${sourcesHtml(it.rk ? esByKey[it.rk] : null, 'eq')}`;
  } else if (mat) {
    h += ` <span class="chr">${esc(mat.intro)}</span></div><div class="meta">Makes ${num(mat.out)} per craft</div>`;
    body += `<h4>Ingredients</h4>${mat.ing.map(([n, q]) => `<div class="rowline"><span class="q">${num(q)}×</span>${link(n)}<span class="k">${chip(kindOf(n).label, kindOf(n).cls)}</span></div>`).join('')}`;
    const rv = rawByMat[name];
    if (rv) {
      const parts = rawCols.map((c, i) => [c, rv[i]]).filter(x => x[1] > 0);
      body += `<h4>Raw materials per 1 unit</h4>${parts.map(([c, q]) => `<div class="rowline"><span class="q">${num(q)}×</span>${link(c)}</div>`).join('')}`;
    }
    body += `<h4>Recipe scroll — where to get it</h4>${sourcesHtml(msByKey['Recipe: ' + name], 'ms')}`;
  } else {
    h += '</div>';
    const eRows = esByKey[name] || esByKey['Recipe: Sealed ' + name.replace(/^Recipe: /, '')], mRows = msByKey[name];
    if (eRows) body += `<h4>Where to get it</h4>${sourcesHtml(eRows, 'eq')}`;
    if (mRows) body += `<h4>Where to get it</h4>${sourcesHtml(mRows, 'ms')}`;
    const target = recipeOf[name];
    if (target) body += `<h4>Teaches</h4>${pills([target.name])}`;
    if (rawCols.includes(name)) {
      const need = D.rawBreakdown.rows.filter(r => r[1][rawCols.indexOf(name)] > 0);
      if (need.length) body += `<h4>Needed (per unit) by</h4>${pills(need.map(r => r[0]))}`;
      else body += '<p class="lead">No recipe in the workbook uses this raw material yet.</p>';
    }
    if (isCurrency(name)) body += '<p class="lead">Crystals and Gemstones are common currency and are not tracked by source.</p>';
    if (!eRows && !mRows && !rawCols.includes(name) && !isCurrency(name) && !target) body += '<p class="lead">No further data in the workbook for this name.</p>';
  }
  body += usedBlock(name);
  return `<button class="iconbtn x" data-close aria-label="Close">Close ✕</button>${h}${body}`;
}

/* ------------------------------------------------------------------ routing */
let current = null, openName = null;
function parseHash() {
  const m = location.hash.replace(/^#\/?/, '').split('/');
  return { view: viewById[m[0]] ? m[0] : 'home', name: m[1] ? decodeURIComponent(m.slice(1).join('/')) : null };
}
function go(view, name) { location.hash = '#/' + view + (name ? '/' + encodeURIComponent(name) : ''); }
function route() {
  const { view, name } = parseHash();
  if (view !== current) {
    current = view; onData = null;
    $('#view').classList.toggle('wide', view === 'raw');
    const el = $('#view'); el.innerHTML = '';
    viewById[view].render(el);
    document.title = viewById[view].label + ' · L2 Reborn Recipe Tracker';
    renderNav(); $('#side').classList.remove('open'); window.scrollTo(0, 0);
  }
  if (name !== openName) { openName = name; name ? showDrawer(name) : hideDrawer(); }
}
function showDrawer(name) {
  const d = $('#drawer'); d.innerHTML = detail(name); d.classList.add('show'); d.setAttribute('aria-hidden', 'false'); d.scrollTop = 0;
  $('#scrim').classList.add('show');
}
function hideDrawer() { $('#drawer').classList.remove('show'); $('#drawer').setAttribute('aria-hidden', 'true'); $('#scrim').classList.remove('show'); }
function closeDrawer() { if (openName) go(current); }
const ICON = {
  home: 'M3 11l9-8 9 8M5 10v10h14V10',
  start: 'M12 3l2.6 5.6 6.1.7-4.5 4.2 1.2 6L12 16.6 6.6 19.5l1.2-6L3.3 9.3l6.1-.7L12 3z',
  materials: 'M9 3h6M10 3v6l-5 9a2 2 0 002 3h10a2 2 0 002-3l-5-9V3',
  raw: 'M6 3h12l3 6-9 12L3 9l3-6zM3 9h18',
  matsrc: 'M12 21s-7-6-7-11a7 7 0 0114 0c0 5-7 11-7 11zM12 12a2 2 0 100-4 2 2 0 000 4z',
  equip: 'M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6l8-3z',
  equipsrc: 'M12 21s-7-6-7-11a7 7 0 0114 0c0 5-7 11-7 11zM12 12a2 2 0 100-4 2 2 0 000 4z',
  notes: 'M6 3h9l3 3v15H6zM9 10h6M9 14h6M9 18h4',
  audit: 'M9 4h6l1 2h3v14H5V6h3l1-2zM9 13l2 2 4-4',
  warehouse: 'M3 8l9-5 9 5v9l-9 5-9-5V8zM3 8l9 5 9-5M12 13v9',
  reqcheck: 'M9 6h11M9 12h11M9 18h11M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2',
  history: 'M12 7v5l3 2M21 12a9 9 0 11-3-6.7M21 4v5h-5',
};
function renderNav() {
  let h = '', g = '';
  VIEWS.filter(v => !v.hidden).forEach(v => {
    if (v.group !== g) { g = v.group; if (g) h += `<div class="navgroup">${esc(g)}</div>`; }
    h += `<a href="#/${v.id}" class="${v.id === current ? 'on' : ''}${v.soon ? ' later' : ''}"><span class="nl"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="${ICON[v.id] || ICON.start}"/></svg>${esc(v.label)}</span>${v.soon ? '<span class="soon">soon</span>' : ''}</a>`;
  });
  $('#nav').innerHTML = `<div class="nav">${h}</div>`;
}

/* --------------------------------------------------------------- global search */
const names = new Map();
function addName(n, rank) { if (n && !names.has(n)) names.set(n, rank); }
D.equipment.forEach(i => addName(i.name, 0)); D.materials.forEach(m => addName(m.name, 1));
rawCols.forEach(n => addName(n, 2));
D.equipSourcing.rows.forEach(r => addName(r[0], 3)); D.matSourcing.rows.forEach(r => addName(r[0], 3));
D.catalogue.forEach(c => addName(c[1], 4));
const nameList = [...names.entries()].map(([n, r]) => ({ n, l: lc(n), r }));
let resIdx = -1, resList = [];
function searchNames(q) {
  q = lc(q.trim()); if (q.length < 2) return [];
  const terms = q.split(/\s+/);
  return nameList.filter(x => terms.every(t => x.l.includes(t)))
    .sort((a, b) => (b.l.startsWith(q) - a.l.startsWith(q)) || a.r - b.r || a.n.localeCompare(b.n)).slice(0, 14);
}
function drawResults() {
  const box = $('#gres'), q = $('#gs').value;
  resList = searchNames(q); resIdx = resList.length ? 0 : -1;
  if (q.trim().length < 2) { box.classList.remove('show'); return; }
  box.innerHTML = resList.length ? resList.map((x, i) => `<button data-open="${esc(x.n)}" class="${i === 0 ? 'cur' : ''}"><span>${esc(x.n)}</span><span class="kind">${esc(kindOf(x.n).label)}</span></button>`).join('') : '<div class="none">No match.</div>';
  box.classList.add('show');
}

/* --------------------------------------------------------------------- setup */
document.addEventListener('click', e => {
  if (e.target.id === 'signout') { B.signOut(); return; }
  const o = e.target.closest('[data-open]');
  if (o) {
    $('#gres').classList.remove('show');
    go(current || 'start', o.dataset.open); return;
  }
  if (e.target.closest('[data-close]') || e.target.id === 'scrim') closeDrawer();
  if (!e.target.closest('.gsearch')) $('#gres').classList.remove('show');
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { $('#gres').classList.remove('show'); if (!$('#modal').hidden) closeModal(true); else closeDrawer(); }
  if (e.key === '/' && !/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName)) { e.preventDefault(); $('#gs').focus(); }
});
const gs = $('#gs');
gs.addEventListener('input', drawResults);
gs.addEventListener('focus', () => gs.value.trim().length >= 2 && drawResults());
gs.addEventListener('keydown', e => {
  const btns = [...document.querySelectorAll('#gres button')];
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault(); if (!btns.length) return;
    resIdx = (resIdx + (e.key === 'ArrowDown' ? 1 : -1) + btns.length) % btns.length;
    btns.forEach((b, i) => b.classList.toggle('cur', i === resIdx)); btns[resIdx].scrollIntoView({ block: 'nearest' });
  }
  if (e.key === 'Enter' && resList[resIdx]) { $('#gres').classList.remove('show'); gs.blur(); go(current || 'start', resList[resIdx].n); }
});
$('#burger').addEventListener('click', () => $('#side').classList.toggle('open'));
$('.brand').addEventListener('click', () => go('home'));
const THEMES = ['gold', 'steel', 'parchment'];
function setTheme(t, keep) {
  if (!THEMES.includes(t)) t = 'gold';
  document.documentElement.dataset.theme = t;
  document.querySelectorAll('#themes [data-theme]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.theme === t)));
  if (!keep) try { localStorage.setItem('l2theme2', t); } catch (e) { /* ignore */ }
}
let saved = null; try { saved = localStorage.getItem('l2theme2'); } catch (e) { /* ignore */ }
setTheme(saved, true);   // nothing chosen yet -> Dark Gold (the old light/dark switch is ignored)
$('#themes').addEventListener('click', e => { const b = e.target.closest('[data-theme]'); if (b) setTheme(b.dataset.theme); });
$('#foot').innerHTML = `Data from workbook <b>${esc(D.meta.version)}</b><br>${esc(D.meta.date)} · exported ${esc(D.meta.exported)}`;
window.addEventListener('hashchange', route);
document.addEventListener('visibilitychange', () => { if (!document.hidden && loaded && (session || B.mode === 'local')) scheduleRefresh(); });
B.onAuth(setSession);
B.init().then(u => { session = u; if (u) { unsub = B.subscribe(scheduleRefresh); maybePresence(); } }, () => {}).then(() => { renderWho(); route(); });
})();
