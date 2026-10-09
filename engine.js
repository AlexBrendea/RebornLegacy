/* L2 Reborn calculator engine - pure logic, no DOM. Works in the browser (window.L2Engine) and in Node (module.exports).
 *
 * Rules (same as the workbook's Requirements Check, plus the fixes agreed for the web edition):
 *  - Each ingredient is taken from stock first. A craftable material (Tier 1-4) that is short is made from its own
 *    ingredients, recursively, using what is left in stock.
 *  - Stock is allocated ingredient by ingredient, so a raw material shared by two shortfalls is never counted twice.
 *  - Batch recipes (e.g. Cord makes 20) are crafted in whole batches; the leftover stays in stock.
 *  - Recipe scroll: D/C-grade recipes are reusable (1 scroll to learn, once). B-grade and above are consumed on every
 *    craft, and the first craft needs one extra scroll to learn the recipe.
 *  - Commit = replace stock with the stock left after all of the above.
 */
(function (root) {
  'use strict';
  const REUSABLE = new Set(['D', 'C']);

  function makeEngine(D) {
    const matByName = {}; D.materials.forEach(m => (matByName[m.name] = m));
    const itemByName = {}; D.equipment.forEach(i => (itemByName[i.name] = i));
    const catByName = {}; D.catalogue.forEach(([cat, name, grade, intro]) => (catByName[name] = { cat, grade, intro }));

    function scrollName(item) {
      const it = typeof item === 'string' ? itemByName[item] : item;
      for (const k of ['Recipe: ' + it.name, 'Recipe: Sealed ' + it.name]) if (catByName[k]) return k;
      return null;
    }
    function scrollNeed(item, qty, learned) {
      const consumed = !REUSABLE.has(item.grade);
      return (consumed ? qty : 0) + (learned ? 0 : 1);
    }
    function kindOf(name) {
      if (matByName[name]) return 'Material';
      const c = catByName[name];
      if (!c) return 'Item';
      if (c.cat === 'Raw Material') return 'Raw material';
      if (c.cat === 'Recipe') return 'Recipe scroll';
      if (c.cat === 'Equipment Part') return 'Equipment part';
      if (c.cat.startsWith('Crystal')) return 'Crystal / Gemstone';
      return c.cat;
    }

    /* Take `amount` of `name` from `stock`, crafting from sub-materials when it is a craftable material.
       Mutates stock and leaves; returns a node describing what happened. */
    function take(name, amount, stock, leaves) {
      const node = { name, needed: amount, fromStock: 0, crafted: 0, crafts: 0, makes: 0, surplus: 0, missing: 0, short: false, children: [] };
      const have = Math.min(stock[name] || 0, amount);
      node.fromStock = have;
      stock[name] = (stock[name] || 0) - have;
      const rest = amount - have;
      const mat = matByName[name];
      if (mat) {
        if (rest > 0) {
          node.crafted = rest;
          node.crafts = Math.ceil(rest / mat.out);
          node.makes = node.crafts * mat.out;
          node.surplus = node.makes - rest;
          for (const [sn, sq] of mat.ing) node.children.push(take(sn, sq * node.crafts, stock, leaves));
          if (node.surplus > 0) stock[name] = (stock[name] || 0) + node.surplus;
          node.short = node.children.some(c => c.short);
        }
      } else {
        const l = leaves[name] || (leaves[name] = { name, required: 0, missing: 0 });
        l.required += amount;
        l.missing += rest;
        node.missing = rest;
        node.short = rest > 0;
      }
      return node;
    }

    /* How many units of craftable material `name` can be made right now from `stock` (not counting units already held). */
    function maxCraftable(name, wanted, stock) {
      const mat = matByName[name];
      if (!mat || wanted <= 0) return 0;
      const ok = k => {
        const s = Object.assign({}, stock), leaves = {};
        const crafts = Math.ceil(k / mat.out);
        return !mat.ing.some(([sn, sq]) => take(sn, sq * crafts, s, leaves).short);
      };
      let lo = 0, hi = wanted;
      while (lo < hi) { const mid = Math.ceil((lo + hi) / 2); if (ok(mid)) lo = mid; else hi = mid - 1; }
      return lo;
    }

    function resolve(itemName, qty, stock0, learned) {
      const item = itemByName[itemName];
      if (!item) throw new Error('Unknown item ' + itemName);
      qty = Math.max(1, Math.floor(qty) || 1);
      const stock = Object.assign({}, stock0), leaves = {}, rows = [];
      const lines = item.ing.map(([n, q]) => ({ name: n, needed: q * qty }));
      const sc = scrollName(item), sNeed = sc ? scrollNeed(item, qty, !!learned) : 0;
      if (sc && sNeed > 0) lines.unshift({ name: sc, needed: sNeed, scroll: true });

      for (const ln of lines) {
        const onHand = stock[ln.name] || 0;                         // what the ingredient has available at this point
        const held = Math.min(onHand, ln.needed);
        const short = ln.needed - held;
        let craftable = 0;
        if (short > 0 && matByName[ln.name]) {
          const snap = Object.assign({}, stock); snap[ln.name] = onHand - held;
          craftable = maxCraftable(ln.name, short, snap);
        }
        const node = take(ln.name, ln.needed, stock, leaves);
        const warehouse = stock0[ln.name] || 0;                     // original quantity, for display
        rows.push({
          name: ln.name, kind: kindOf(ln.name), isScroll: !!ln.scroll, needed: ln.needed, warehouse,
          short: Math.max(0, ln.needed - warehouse), craftable, balance: warehouse + craftable - ln.needed,
          ok: !node.short, node,
        });
      }
      const shopping = Object.values(leaves).filter(l => l.missing > 0).map(l => ({
        name: l.name, kind: kindOf(l.name), required: l.required, warehouse: stock0[l.name] || 0, missing: l.missing,
      }));
      const changes = [];
      for (const n of new Set([...Object.keys(stock0), ...Object.keys(stock)])) {
        const a = stock0[n] || 0, b = stock[n] || 0;
        if (a !== b) changes.push({ name: n, before: a, after: b });
      }
      return { item, qty, scroll: sc ? { name: sc, needed: sNeed, consumed: !REUSABLE.has(item.grade), learned: !!learned } : null,
        rows, shopping, ok: rows.every(r => r.ok), finalStock: stock, changes };
    }

    return { resolve, scrollName, scrollNeed, kindOf, maxCraftable, matByName, itemByName, catByName, REUSABLE };
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = makeEngine;
  root.L2Engine = makeEngine;
})(typeof window !== 'undefined' ? window : globalThis);
