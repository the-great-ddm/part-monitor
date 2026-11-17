/**
 * Daya-LAE Part Monitor — parameter rumus per supplier & penyesuaian per item (fungsi murni, tanpa DOM/fetch).
 * Aturan: docs/FORMULAS.md §6. Diatur dari menu "Parameter", disimpan di Sheets (tab suppliers, item_rules, param_log).
 *
 * settings = {
 *   suppliers: [{ id, name, city, category: 'PART'|'BAHAN', prefixes: ['SAN-OIL-KIX-', …],
 *                 lead, cycle, safety,            // hari: lead time + siklus order + safety stock → Stock Level = Σ/30
 *                 level_round: 'UP1'|'NONE',      // pembulatan Stock Level
 *                 mip_round: 'ROUND0'|'ROUNDUP1', // pembulatan MIP
 *                 factor,                         // pengali MIP dalam % (100 = tanpa perubahan)
 *                 multiple,                       // kelipatan order (0 = tidak dibulatkan)
 *                 tol_days }],                    // toleransi overstock dalam hari pemakaian
 *   rules: [{ key,            // kode item persis, atau awalan diakhiri '*' (mis. 'D90000-*')
 *             supplier, mip, level, factor, multiple, reason, by, at }],   // kosong = ikut supplier
 * }
 * Supplier tanpa awalan = supplier bawaan kategorinya (semua item yang tidak cocok awalan mana pun).
 * Tanpa perubahan, DEFAULT_SETTINGS menghasilkan angka yang sama persis dengan rumus Excel (FORMULAS.md §1).
 */
(function (root) {
  'use strict';

  const MONTHS = 6; // jendela demand mengikuti tarikan ADMS (6 bulan bergulir), tidak diatur per supplier

  const sup = (id, name, category, prefixes, lead, cycle, safety, extra = {}) => ({
    id, name, city: '', category, prefixes, lead, cycle, safety,
    level_round: category === 'PART' ? 'UP1' : 'NONE',
    mip_round: category === 'PART' ? 'ROUND0' : 'ROUNDUP1',
    factor: 100, multiple: category === 'PART' ? 1 : 0, tol_days: 0, ...extra,
  });

  /** Sama dengan rumus Excel Bagas: Part 15+15+14, Bahan 7+7+15, oli KIXX 15+15+15. Bahan dikelompokkan per awalan kode. */
  const DEFAULT_SETTINGS = {
    suppliers: [
      sup('ATPM', 'ATPM Daihatsu', 'PART', [], 15, 15, 14),
      sup('BAHAN', 'Bahan lainnya', 'BAHAN', [], 7, 7, 15),
      sup('KIXX', 'Oli KIXX', 'BAHAN', ['SAN-OIL-KIX-'], 15, 15, 15),
      sup('SAN', 'SAN', 'BAHAN', ['SAN-'], 7, 7, 15),
      sup('PRM', 'PRM', 'BAHAN', ['PRM-'], 7, 7, 15),
      sup('OTG', 'OTG', 'BAHAN', ['OTG-'], 7, 7, 15),
      sup('ICG', 'ICG (aki)', 'BAHAN', ['ICG-'], 7, 7, 15),
      sup('SJE', 'SJE', 'BAHAN', ['SJE-'], 7, 7, 15),
      sup('CPU', 'CPU', 'BAHAN', ['CPU-'], 7, 7, 15),
      sup('GTA', 'GTA', 'BAHAN', ['GTA'], 7, 7, 15),
    ],
    rules: [],
  };

  const clone = (x) => JSON.parse(JSON.stringify(x));
  const numOrNull = (v) => (v === null || v === undefined || v === '' || !Number.isFinite(Number(v)) ? null : Number(v));
  const r12 = (x) => Number.parseFloat(Number(x).toPrecision(12));

  /** Stock Level supplier (bulan), dengan pembulatan yang sama seperti calc.js. */
  function supplierLevel(s) {
    const raw = ((Number(s.lead) || 0) + (Number(s.cycle) || 0) + (Number(s.safety) || 0)) / 30;
    return s.level_round === 'UP1' ? Math.ceil(r12(raw * 10)) / 10 : raw;
  }

  /** Aturan item yang berlaku: kode persis dulu, lalu awalan* terpanjang. */
  function ruleFor(code, settings) {
    let best = null, bestLen = -1;
    for (const r of settings.rules || []) {
      const k = String(r.key || '').trim().toUpperCase();
      if (!k) continue;
      if (k === code) return r;
      if (k.endsWith('*') && code.startsWith(k.slice(0, -1)) && k.length > bestLen) { best = r; bestLen = k.length; }
    }
    return best;
  }

  /** Supplier item: dari aturan item, lalu awalan kode terpanjang dalam kategorinya, lalu supplier bawaan kategori. */
  function supplierFor(item, settings, rule = ruleFor(item.item_code, settings)) {
    const list = settings.suppliers || [];
    if (rule && rule.supplier) {
      const s = list.find((x) => x.id === rule.supplier);
      if (s) return s;
    }
    let best = null, bestLen = -1;
    for (const s of list) {
      if (s.category !== item.category) continue;
      for (const p of s.prefixes || []) {
        const pre = String(p).trim().toUpperCase();
        if (pre && item.item_code.startsWith(pre) && pre.length > bestLen) { best = s; bestLen = pre.length; }
      }
    }
    return best || list.find((s) => s.category === item.category && !(s.prefixes || []).some((p) => String(p).trim()))
      || list.find((s) => s.category === item.category) || null;
  }

  /**
   * Input calc.calcItem untuk satu item: { params: {PART|BAHAN: {...}}, fields: {level_override, mip_override, mip_factor, moq}, supplier, rule }.
   * Tanpa supplier yang cocok → parameter bawaan calc.js.
   */
  function resolve(item, settings, calcDefaults) {
    const rule = ruleFor(item.item_code, settings);
    const s = supplierFor(item, settings, rule);
    const base = calcDefaults[item.category];
    const p = s ? {
      months: MONTHS,
      days: [Number(s.lead) || 0, Number(s.cycle) || 0, Number(s.safety) || 0],
      levelRound: s.level_round === 'UP1' ? 1 : null,
      mipRound: s.mip_round === 'ROUND0' ? 'ROUND0' : 'ROUNDUP1',
      moq: Number(s.multiple) || 0,
      tolDays: Number(s.tol_days) || 0,
    } : { ...base, tolDays: 0 };
    const fs = numOrNull(s && s.factor), fr = numOrNull(rule && rule.factor);
    const factor = (fs == null ? 100 : fs) * (fr == null ? 100 : fr) / 10000;
    const fields = {
      level_override: numOrNull(rule && rule.level),
      mip_override: numOrNull(rule && rule.mip),
      mip_factor: factor === 1 ? null : factor,
      moq: numOrNull(rule && rule.multiple) != null ? Number(rule.multiple) : undefined,
    };
    return { params: { [item.category]: p }, fields, supplier: s, rule };
  }

  /** Periksa isian sebelum disimpan → daftar pesan (kosong = aman). */
  function validate(settings) {
    const errs = [];
    const ids = new Set();
    for (const cat of ['PART', 'BAHAN']) {
      if (!settings.suppliers.some((s) => s.category === cat && !(s.prefixes || []).some((p) => String(p).trim()))) {
        errs.push(`Harus ada satu supplier ${cat === 'PART' ? 'Part' : 'Bahan'} tanpa awalan kode (supplier bawaan untuk item yang tidak cocok awalan mana pun).`);
      }
    }
    const seen = {};
    for (const s of settings.suppliers) {
      const name = s.name || s.id;
      if (!String(s.name || '').trim()) errs.push('Ada supplier tanpa nama.');
      if (ids.has(s.id)) errs.push(`Kode supplier ${s.id} dobel.`);
      ids.add(s.id);
      for (const f of ['lead', 'cycle', 'safety', 'factor', 'multiple', 'tol_days']) {
        const v = Number(s[f]);
        if (!Number.isFinite(v) || v < 0) errs.push(`${name}: isian ${LABEL[f]} harus angka ≥ 0.`);
      }
      if ((Number(s.lead) || 0) + (Number(s.cycle) || 0) + (Number(s.safety) || 0) <= 0) errs.push(`${name}: lead time + siklus + safety stock tidak boleh 0.`);
      if (Number(s.factor) > 500) errs.push(`${name}: pengali MIP maksimal 500%.`);
      for (const p of s.prefixes || []) {
        const k = s.category + ':' + String(p).trim().toUpperCase();
        if (!String(p).trim()) continue;
        if (seen[k]) errs.push(`Awalan "${p}" dipakai dua supplier (${seen[k]} dan ${name}).`);
        seen[k] = name;
      }
    }
    const keys = new Set();
    for (const r of settings.rules) {
      const k = String(r.key || '').trim().toUpperCase();
      if (!k) { errs.push('Ada penyesuaian item tanpa kode.'); continue; }
      if (keys.has(k)) errs.push(`Penyesuaian untuk ${k} dobel.`);
      keys.add(k);
      if (r.supplier && !ids.has(r.supplier)) errs.push(`${k}: supplier tidak ditemukan.`);
      for (const f of ['mip', 'level', 'factor', 'multiple']) {
        const v = numOrNull(r[f]);
        if (r[f] !== '' && r[f] != null && (v == null || v < 0)) errs.push(`${k}: ${LABEL[f]} harus angka ≥ 0.`);
      }
      if (numOrNull(r.level) > 12) errs.push(`${k}: Stock Level manual maksimal 12 bulan.`);
      if (numOrNull(r.factor) > 500) errs.push(`${k}: pengali MIP maksimal 500%.`);
      const any = r.supplier || ['mip', 'level', 'factor', 'multiple'].some((f) => numOrNull(r[f]) != null);
      if (!any) errs.push(`${k}: belum ada yang diatur.`);
      if (any && !String(r.reason || '').trim()) errs.push(`${k}: alasan wajib diisi (misal: instruksi management).`);
    }
    return errs;
  }

  const LABEL = {
    name: 'nama', city: 'kota', category: 'kategori', prefixes: 'awalan kode', lead: 'lead time', cycle: 'siklus order', safety: 'safety stock',
    level_round: 'pembulatan Stock Level', mip_round: 'pembulatan MIP', factor: 'pengali MIP', multiple: 'kelipatan order',
    tol_days: 'toleransi overstock', supplier: 'supplier', mip: 'MIP manual', level: 'Stock Level manual', reason: 'alasan',
  };
  const show = (f, v) => {
    if (v === null || v === undefined || v === '') return '–';
    if (Array.isArray(v)) return v.length ? v.join(', ') : '–';
    if (f === 'factor') return v + '%';
    if (['lead', 'cycle', 'safety', 'tol_days'].includes(f)) return v + ' hari';
    if (f === 'level') return v + ' bln';
    return String(v);
  };

  /** Ringkasan perubahan (untuk riwayat & konfirmasi) → daftar kalimat. */
  function diff(before, after) {
    const out = [];
    const S_FIELDS = ['name', 'city', 'category', 'prefixes', 'lead', 'cycle', 'safety', 'level_round', 'mip_round', 'factor', 'multiple', 'tol_days'];
    const R_FIELDS = ['supplier', 'mip', 'level', 'factor', 'multiple', 'reason'];
    const same = (a, b) => JSON.stringify(a ?? '') === JSON.stringify(b ?? '');
    const bs = Object.fromEntries(before.suppliers.map((s) => [s.id, s]));
    const as = Object.fromEntries(after.suppliers.map((s) => [s.id, s]));
    for (const s of after.suppliers) {
      const o = bs[s.id];
      if (!o) { out.push(`Supplier baru: ${s.name} (${s.category}, ${s.lead}+${s.cycle}+${s.safety} hari)`); continue; }
      const ch = S_FIELDS.filter((f) => !same(o[f], s[f])).map((f) => `${LABEL[f]} ${show(f, o[f])} → ${show(f, s[f])}`);
      if (ch.length) out.push(`Supplier ${s.name}: ${ch.join(', ')}`);
    }
    for (const s of before.suppliers) if (!as[s.id]) out.push(`Supplier dihapus: ${s.name}`);
    const key = (r) => String(r.key || '').trim().toUpperCase();
    const br = Object.fromEntries(before.rules.map((r) => [key(r), r]));
    const ar = Object.fromEntries(after.rules.map((r) => [key(r), r]));
    const supName = (id) => (id ? (as[id] || bs[id] || { name: id }).name : '');
    const val = (f, r) => (f === 'supplier' ? supName(r.supplier) : r[f]);
    for (const r of after.rules) {
      const o = br[key(r)];
      const set = R_FIELDS.filter((f) => f !== 'reason' && val(f, r) !== '' && val(f, r) != null).map((f) => `${LABEL[f]} ${show(f, val(f, r))}`);
      if (!o) { out.push(`Item ${key(r)}: ${set.join(', ')} (alasan: ${r.reason || '–'})`); continue; }
      const ch = R_FIELDS.filter((f) => !same(val(f, o), val(f, r))).map((f) => `${LABEL[f]} ${show(f, val(f, o))} → ${show(f, val(f, r))}`);
      if (ch.length) out.push(`Item ${key(r)}: ${ch.join(', ')}`);
    }
    for (const r of before.rules) if (!ar[key(r)]) out.push(`Penyesuaian item ${key(r)} dihapus`);
    return out;
  }

  // ---- baris Sheets ----
  const SUP_COLS = ['id', 'name', 'city', 'category', 'prefixes', 'lead', 'cycle', 'safety', 'level_round', 'mip_round', 'factor', 'multiple', 'tol_days'];
  const RULE_COLS = ['key', 'supplier', 'mip', 'level', 'factor', 'multiple', 'reason', 'by', 'at'];
  const SUP_NUM = ['lead', 'cycle', 'safety', 'factor', 'multiple', 'tol_days'];
  const RULE_NUM = ['mip', 'level', 'factor', 'multiple'];

  function toRows(settings) {
    return {
      suppliers: settings.suppliers.map((s) => Object.fromEntries(SUP_COLS.map((c) => [c, c === 'prefixes' ? (s.prefixes || []).join(',') : s[c] ?? '']))),
      rules: settings.rules.map((r) => Object.fromEntries(RULE_COLS.map((c) => [c, c === 'key' ? String(r.key || '').trim().toUpperCase() : r[c] ?? '']))),
    };
  }

  /** Dari server → settings. Tab kosong / belum ada → null (pakai bawaan). */
  function fromRows(rows) {
    if (!rows || !Array.isArray(rows.suppliers) || !rows.suppliers.length) return null;
    return {
      suppliers: rows.suppliers.map((r) => {
        const s = {};
        for (const c of SUP_COLS) s[c] = SUP_NUM.includes(c) ? Number(r[c]) || 0 : String(r[c] ?? '');
        s.prefixes = String(r.prefixes || '').split(',').map((p) => p.trim().toUpperCase()).filter(Boolean);
        s.category = s.category === 'PART' ? 'PART' : 'BAHAN';
        if (r.factor === '' || r.factor == null) s.factor = 100;
        return s;
      }),
      rules: (rows.rules || []).map((r) => {
        const o = {};
        for (const c of RULE_COLS) o[c] = RULE_NUM.includes(c) ? numOrNull(r[c]) : String(r[c] ?? '');
        return o;
      }),
    };
  }

  const api = { DEFAULT_SETTINGS, MONTHS, LABEL, clone, supplierLevel, ruleFor, supplierFor, resolve, validate, diff, toRows, fromRows, SUP_COLS, RULE_COLS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LAEPart = Object.assign(root.LAEPart || {}, { params: api });
})(typeof window !== 'undefined' ? window : globalThis);
