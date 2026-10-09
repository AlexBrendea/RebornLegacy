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
const head = (title, lead) => `<h1>${esc(title)}</h1>${lead ? `<p class="lead">${esc(lead)}</p>` : ''}`;
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
    <div class="banner"><b>Web edition — step 1: reference tabs.</b> Everything below is read from the workbook (${esc(D.meta.version)}). The live clan Warehouse, the Requirements Check calculator and the History log will appear in the three greyed-out tabs once the shared database is connected.</div>
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
  el.innerHTML = head('Material Recipes', 'What each Tier 1–4 material is made from. Tier 1 uses raw materials only; higher tiers build on lower ones.') + '<div id="t"></div>';
  makeTable($('#t', el), {
    rows: D.materials, noun: 'materials', placeholder: 'Search material or ingredient…',
    search: m => m.name + ' ' + m.ing.map(i => i[0]).join(' '),
    stripe: m => tierColor(m.tier),
    filters: [
      { id: 'tier', label: 'Tier', options: ['1', '2', '3', '4'], test: (m, v) => String(m.tier) === v },
      { id: 'intro', label: 'Introduced', options: uniq(D.materials.map(m => m.intro)).sort(byOrder(CHRON)), test: (m, v) => m.intro === v },
    ],
    cols: [
      { key: 'name', label: 'Material', get: m => m.name, html: m => link(m.name) },
      { key: 'tier', label: 'Tier', get: m => m.tier, html: m => chip('Tier ' + m.tier, 't-' + m.tier) },
      { key: 'intro', label: 'Introduced', get: m => CHRON.indexOf(m.intro), html: m => `<span class="chr">${esc(m.intro)}</span>` },
      { key: 'out', label: 'Output qty', cls: 'num', get: m => m.out, html: m => num(m.out) },
      { key: 'ing', label: 'Ingredients', nosort: true, get: m => '', html: m => ingList(m.ing) },
    ],
  }, st('materials'));
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
  }
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

const VIEWS = [
  { id: 'start', label: 'Start Here', group: 'Reference', render: startView },
  { id: 'materials', label: 'Material Recipes', group: 'Reference', render: materialsView },
  { id: 'raw', label: 'Raw Material Breakdown', group: 'Reference', render: rawView },
  { id: 'matsrc', label: 'Material Recipe Sourcing', group: 'Reference', render: matSourcingView },
  { id: 'equip', label: 'Equipment Recipes', group: 'Reference', render: equipmentView },
  { id: 'equipsrc', label: 'Equipment Sourcing', group: 'Reference', render: equipSourcingView },
  { id: 'notes', label: 'Notes', group: 'Reference', render: notesView },
  { id: 'audit', label: 'Audit Log', group: 'Reference', render: auditView },
  { id: 'warehouse', label: 'Warehouse', group: 'Clan (live)', soon: true, render: soonView('Warehouse', 'The shared clan stock.', ['Current quantity of every material, recipe scroll, part and crystal, filtered by category and grade.', 'Add and Remove buttons for deposits and withdrawals — both of you see the same numbers.']) },
  { id: 'reqcheck', label: 'Requirements Check', group: 'Clan (live)', soon: true, render: soonView('Requirements Check', 'The crafting calculator.', ['Pick a grade, an item and a quantity: see Needed vs Warehouse vs Balance.', 'Missing ingredients are broken down into the sub-materials you still need.', 'Commit (enabled when every balance is green) subtracts the materials from the shared warehouse in one step.']) },
  { id: 'history', label: 'History', group: 'Clan (live)', soon: true, render: soonView('History', 'Who changed what, and when.', ['Every deposit, withdrawal and commit with the person and the time.']) },
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
    const eRows = esByKey[name], mRows = msByKey[name];
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
  return { view: viewById[m[0]] ? m[0] : 'start', name: m[1] ? decodeURIComponent(m.slice(1).join('/')) : null };
}
function go(view, name) { location.hash = '#/' + view + (name ? '/' + encodeURIComponent(name) : ''); }
function route() {
  const { view, name } = parseHash();
  if (view !== current) {
    current = view;
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
function renderNav() {
  let h = '', g = '';
  VIEWS.forEach(v => {
    if (v.group !== g) { g = v.group; h += `<div class="navgroup">${esc(g)}</div>`; }
    h += `<a href="#/${v.id}" class="${v.id === current ? 'on' : ''}${v.soon ? ' later' : ''}">${esc(v.label)}${v.soon ? '<span class="soon">soon</span>' : ''}</a>`;
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
  const o = e.target.closest('[data-open]');
  if (o) {
    $('#gres').classList.remove('show');
    go(current || 'start', o.dataset.open); return;
  }
  if (e.target.closest('[data-close]') || e.target.id === 'scrim') closeDrawer();
  if (!e.target.closest('.gsearch')) $('#gres').classList.remove('show');
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { $('#gres').classList.remove('show'); closeDrawer(); }
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
function setTheme(t) { document.documentElement.dataset.theme = t; try { localStorage.setItem('l2theme', t); } catch (e) { /* ignore */ } }
let saved = null; try { saved = localStorage.getItem('l2theme'); } catch (e) { /* ignore */ }
setTheme(saved || (matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark'));
$('#theme').addEventListener('click', () => setTheme(document.documentElement.dataset.theme === 'light' ? 'dark' : 'light'));
$('#foot').innerHTML = `Data from workbook <b>${esc(D.meta.version)}</b><br>${esc(D.meta.date)} · exported ${esc(D.meta.exported)}`;
window.addEventListener('hashchange', route);
route();
})();
