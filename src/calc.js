/**
 * Daya-LAE Part Monitor — mesin hitung (fungsi murni, tanpa DOM/fetch).
 * Rumus mengikuti Excel referensi Bagas (lihat docs/FORMULAS.md, status LOCKED 16 Nov 2026).
 *
 *   MAD          = Demand(6 bln) / 6
 *   Stock Level  = (hari1 + hari2 + hari3) / 30
 *                  PART  : ROUNDUP(…, 1)       → 15+15+14 hari = 1.5
 *                  BAHAN : tanpa pembulatan    → 7+7+15 = 0.9667 ; KIXX 15+15+15 = 1.5
 *   MIP          = PART  : ROUND(StockLevel × MAD, 0)
 *                  BAHAN : ROUNDUP(StockLevel × MAD, 1)
 *   SOQ          = MIP − Stock − On Order   (boleh negatif = kelebihan stok; On Order = PO yang belum datang, default 0)
 *   SOQ Order    = max(0, SOQ)        (yang benar-benar dipesan; dibulatkan ke atas ke kelipatan order / MOQ bila > 0)
 *   Stock Rasio  = Stock / MAD        (hanya PART; MAD 0 → null)
 *
 * Penyesuaian dari menu Parameter (docs/FORMULAS.md §6, semuanya opsional; tanpa itu hasil = rumus Excel di atas):
 *   params.tolDays      toleransi overstock (hari pemakaian): Normal selama Stock ≤ MIP + tolDays/30 × MAD
 *   item.level_override Stock Level manual (bulan) menggantikan (hari1+hari2+hari3)/30
 *   item.mip_factor     pengali MIP (1 = 100%): MIP = pembulatan(Stock Level × MAD × pengali)
 *   item.mip_override   MIP manual (angka pasti) menggantikan hasil rumus
 */
(function (root) {
  'use strict';

  // Hilangkan noise floating point sebelum pembulatan ala Excel (Excel bekerja di 15 digit signifikan).
  const clean = (x) => Number.parseFloat(Number(x).toPrecision(12));

  /** Excel ROUND: half away from zero */
  function xlRound(x, d = 0) {
    const f = 10 ** d, v = clean(x * f);
    return Math.sign(v) * Math.round(Math.abs(v)) / f;
  }
  /** Excel ROUNDUP: menjauhi nol */
  function xlRoundUp(x, d = 0) {
    const f = 10 ** d, v = clean(x * f);
    return Math.sign(v) * Math.ceil(Math.abs(v)) / f;
  }

  const DEFAULT_PARAMS = {
    PART:  { months: 6, days: [15, 15, 14], levelRound: 1,    mipRound: 'ROUND0',   moq: 1 },
    BAHAN: { months: 6, days: [7, 7, 15],   levelRound: null, mipRound: 'ROUNDUP1', moq: 0 }, // 0 = order tidak dibulatkan
  };

  /** Agregasi demand per item dari baris transaksi. method: 'SUM' (qty) | 'COUNT' (jumlah baris, seperti pivot Excel Part saat ini) */
  function aggregateDemand(rows, { fromPeriod, toPeriod, method = 'SUM' } = {}) {
    const out = {};
    for (const r of rows) {
      const p = r.period; // 'YYYY-MM'
      if (fromPeriod && p < fromPeriod) continue;
      if (toPeriod && p > toPeriod) continue;
      const k = String(r.item_code).trim().toUpperCase();
      out[k] = (out[k] || 0) + (method === 'COUNT' ? 1 : Number(r.qty) || 0);
    }
    return out;
  }

  function stockLevel(days, levelRound) {
    const raw = days.reduce((a, d) => a + d / 30, 0);
    return levelRound == null ? raw : xlRoundUp(raw, levelRound);
  }

  function mip(level, mad, mode) {
    const x = level * mad;
    if (mode === 'ROUND0') return xlRound(x, 0);
    if (mode === 'ROUNDUP1') return xlRoundUp(x, 1);
    throw new Error('mipRound tidak dikenal: ' + mode);
  }

  /**
   * DEAD      : tidak ada demand 6 bln, masih ada stok (modal tertahan)
   * NO_MOVE   : tidak ada demand, tidak ada stok
   * NON_STOCK : ada demand tapi sangat jarang → MIP 0, stok 0 (cukup dipesan saat ada kebutuhan / SOP)
   * STOCKOUT  : stok 0 padahal MIP > 0, dan PO yang berjalan belum menutup kebutuhan → prioritas order
   * ORDER     : stok ada tapi di bawah MIP, dan PO yang berjalan belum menutup kebutuhan
   * INCOMING  : stok di bawah MIP tetapi sudah tertutup PO yang belum datang → tunggu, jangan order lagi
   * OVERSTOCK : stok di atas MIP
   * Tanpa On Order (0) hasilnya sama persis dengan aturan sebelum ada fitur PO.
   */
  function status({ stock, mad, mip, soq, tol = 0 }) {
    // MIP manual > 0 untuk item tanpa demand: tetap dihitung seperti item biasa (management ingin item ini distok)
    if (mad === 0 && !(mip > 0)) return stock > 0 ? 'DEAD' : 'NO_MOVE';
    if (stock <= 0 && mip <= 0) return 'NON_STOCK';
    if (soq > 0) return stock <= 0 ? 'STOCKOUT' : 'ORDER';
    if (stock < mip) return 'INCOMING';
    return stock > clean(mip + tol) ? 'OVERSTOCK' : 'NORMAL';
  }

  /**
   * Hitung satu item.
   * item: { item_code, category: 'PART'|'BAHAN', stock, demand, on_order?, days?, moq? }
   */
  function calcItem(item, params = DEFAULT_PARAMS) {
    const p = params[item.category];
    if (!p) throw new Error('Kategori harus PART atau BAHAN: ' + item.item_code);
    const stock = Math.max(0, Number(item.stock) || 0);
    const demand = Number(item.demand) || 0;
    const mad = demand / p.months;
    const days = item.days || p.days;
    const num = (v) => (v === null || v === undefined || v === '' || !Number.isFinite(Number(v)) ? null : Number(v));
    const levelOv = num(item.level_override), mipOv = num(item.mip_override), factor = num(item.mip_factor);
    const level = levelOv != null ? levelOv : stockLevel(days, p.levelRound);
    const m = mipOv != null ? Math.max(0, mipOv) : mip(level, factor != null ? mad * factor : mad, p.mipRound);
    const tol = (Number(p.tolDays) || 0) / 30 * mad;
    const onOrder = Math.max(0, Number(item.on_order) || 0);
    const soq = clean(m - stock - onOrder);
    const moq = Number(item.moq || p.moq) || 0;
    let soqOrder = Math.max(0, soq);
    if (moq > 0 && soqOrder > 0) soqOrder = clean(Math.ceil(clean(soqOrder / moq)) * moq);
    return {
      ...item,
      stock,
      on_order: onOrder,
      demand,
      mad,
      stock_ratio: mad > 0 ? stock / mad : null,
      stock_level: level,
      mip: m,
      soq,
      soq_order: soqOrder,
      status: status({ stock, mad, mip: m, soq, tol }),
      warnings: Number(item.stock) < 0 ? ['Stock negatif dari ADMS dianggap 0'] : [],
    };
  }

  const calcAll = (items, params = DEFAULT_PARAMS) => items.map((it) => calcItem(it, params));

  const api = { xlRound, xlRoundUp, aggregateDemand, stockLevel, mip, status, calcItem, calcAll, DEFAULT_PARAMS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LAEPart = Object.assign(root.LAEPart || {}, { calc: api });
})(typeof window !== 'undefined' ? window : globalThis);
