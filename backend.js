/* Storage layer for the clan warehouse.
   - Shared mode: Supabase (needs config.js + the supabase-js script). Login required.
   - Test mode: this browser only (used when config.js is missing, or when the address ends in ?local=1).
   Both expose the same functions, so the screens do not care which one is running. */
(function () {
'use strict';
const cfg = window.L2_CONFIG || {};
const wantLocal = /[?&]local=1/.test(location.search);
const shared = !!(cfg.url && cfg.anonKey && window.supabase && window.supabase.createClient && !wantLocal);

/* ------------------------------------------------------------ shared (Supabase) */
function makeShared() {
  const client = window.supabase.createClient(cfg.url, cfg.anonKey, { auth: { persistSession: true, autoRefreshToken: true } });
  let user = null; const listeners = [];
  const toUser = s => (s && s.user ? { email: String(s.user.email || '').toLowerCase() } : null);
  const fail = r => { if (r.error) { const e = new Error(r.error.message || 'Request failed'); e.code = r.error.code; throw e; } return r.data; };
  return {
    mode: 'shared', canDemo: false,
    async init() {
      const { data } = await client.auth.getSession(); user = toUser(data && data.session);
      client.auth.onAuthStateChange((ev, s) => {
        const nu = toUser(s);
        if ((nu && nu.email) !== (user && user.email)) { user = nu; setTimeout(() => listeners.forEach(f => f(user)), 0); }
      });
      return user;
    },
    onAuth(f) { listeners.push(f); },
    async signIn(email, password) { const r = await client.auth.signInWithPassword({ email, password }); return { error: r.error ? r.error.message : null }; },
    async signOut() { await client.auth.signOut(); },
    async load() {
      const ok = fail(await client.rpc('is_allowed'));
      if (!ok) { const e = new Error('not-allowed'); e.notAllowed = true; throw e; }
      const [s, l] = await Promise.all([
        client.from('stock').select('item,qty,updated_at,updated_by').range(0, 4999),
        client.from('learned').select('item,learned_by,learned_at').range(0, 4999)]);
      const stock = {}, meta = {}, learned = {};
      (fail(s) || []).forEach(r => { stock[r.item] = r.qty; meta[r.item] = { by: r.updated_by, at: r.updated_at }; });
      (fail(l) || []).forEach(r => { learned[r.item] = true; });
      return { stock, learned, meta };
    },
    async apply(changes, action, label, learn) {
      const r = await client.rpc('apply_changes', { p_changes: changes, p_action: action || 'set', p_label: label || null, p_learn: learn || null });
      return fail(r);
    },
    async setLearned(item, flag) { fail(await client.rpc('set_learned', { p_item: item, p_learned: !!flag })); },
    async history(limit) {
      const r = await client.from('history').select('id,at,by,action,item,old_qty,new_qty,batch,label').order('id', { ascending: false }).limit(limit || 1000);
      return fail(r) || [];
    },
    subscribe(cb) {
      let ch = null;
      try {
        ch = client.channel('l2-live');
        ['stock', 'learned', 'history'].forEach(t => ch.on('postgres_changes', { event: '*', schema: 'public', table: t }, cb));
        ch.subscribe();
      } catch (e) { /* live updates are a bonus; the page also refreshes when you come back to it */ }
      return () => { try { ch && client.removeChannel(ch); } catch (e) { /* ignore */ } };
    },
    /* Who is online right now (Supabase Presence). Only the part of the e-mail before the @ is shared.
       Uses a private channel when the database allows it (supabase/presence.sql); otherwise a public one. */
    presence(cb) {
      if (!user) return () => {};
      const handle = String(user.email).split('@')[0];
      let ch = null, dead = false;
      const open = priv => {
        const c = client.channel('l2-online', { config: Object.assign({ presence: { key: handle } }, priv ? { private: true } : {}) });
        c.on('presence', { event: 'sync' }, () => { try { cb(Object.keys(c.presenceState())); } catch (e) { /* ignore */ } });
        c.subscribe(async status => {
          if (dead) return;
          if (status === 'SUBSCRIBED') { try { await c.track({ at: new Date().toISOString() }); } catch (e) { /* ignore */ } }
          else if (priv && (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT')) { try { client.removeChannel(c); } catch (e) { /* ignore */ } if (!dead) ch = open(false); }
        });
        return c;
      };
      try { ch = open(true); } catch (e) { /* online status is a bonus */ }
      return () => { dead = true; try { ch && client.removeChannel(ch); } catch (e) { /* ignore */ } };
    },
  };
}

/* ------------------------------------------------------------ test (this browser) */
function makeLocal() {
  const KEY = 'l2.testWarehouse.v2';
  let db = { stock: {}, learned: {}, meta: {}, hist: [], seq: 0 };
  try { const sv = JSON.parse(localStorage.getItem(KEY) || 'null'); if (sv && sv.stock) db = Object.assign(db, sv); } catch (e) { /* ignore */ }
  // v1 test warehouses (step 2) are carried over once
  try { if (!db.seq && !Object.keys(db.stock).length) { const o = JSON.parse(localStorage.getItem('l2.testWarehouse.v1') || 'null'); if (o && o.stock) { db.stock = o.stock; db.learned = o.learned || {}; } } } catch (e) { /* ignore */ }
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(db)); } catch (e) { /* ignore */ } };
  const ME = 'test';
  const log = (action, item, o, n, batch, label) => { db.hist.unshift({ id: ++db.seq, at: new Date().toISOString(), by: ME, action, item, old_qty: o, new_qty: n, batch, label: label || null }); db.hist.length = Math.min(db.hist.length, 500); };
  const snap = () => ({ stock: Object.assign({}, db.stock), learned: Object.assign({}, db.learned), meta: JSON.parse(JSON.stringify(db.meta)) });
  const subs = [];
  const ping = () => subs.forEach(f => f());
  return {
    mode: 'local', canDemo: true,
    async init() { return { email: ME }; },
    onAuth() {}, async signIn() { return { error: null }; }, async signOut() {},
    async load() { return snap(); },
    async apply(changes, action, label, learn) {
      const conflicts = [];
      changes.forEach(c => { const cur = db.stock[c.item] || 0; if (c.expected != null && c.expected !== cur) conflicts.push({ item: c.item, expected: c.expected, current: cur, by: (db.meta[c.item] || {}).by, at: (db.meta[c.item] || {}).at, new: c.new }); });
      if (conflicts.length) return { ok: false, conflicts };
      const batch = 'b' + (db.seq + 1); let changed = 0;
      changes.forEach(c => { const cur = db.stock[c.item] || 0; if (c.new === cur) return; db.stock[c.item] = c.new; db.meta[c.item] = { by: ME, at: new Date().toISOString() }; log(action || 'set', c.item, cur, c.new, batch, label); changed++; });
      if (learn && !db.learned[learn]) { db.learned[learn] = true; log('learn', learn, null, null, batch, label); }
      save(); ping(); return { ok: true, changed };
    },
    async setLearned(item, flag) { if (flag && !db.learned[item]) { db.learned[item] = true; log('learn', item, null, null, 'b' + (db.seq + 1)); } else if (!flag && db.learned[item]) { delete db.learned[item]; log('unlearn', item, null, null, 'b' + (db.seq + 1)); } save(); ping(); },
    async history(limit) { return db.hist.slice(0, limit || 1000); },
    subscribe(cb) { subs.push(cb); return () => { const i = subs.indexOf(cb); if (i >= 0) subs.splice(i, 1); }; },
    /* test-mode only helpers */
    async replaceAll(stock, learned) { db.stock = stock || {}; db.learned = learned || {}; db.meta = {}; save(); ping(); },
    async dump() { return { stock: db.stock, learned: db.learned }; },
  };
}

window.L2Backend = shared ? makeShared() : makeLocal();
})();
