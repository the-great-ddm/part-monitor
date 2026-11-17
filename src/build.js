/**
 * Daya-LAE Part Monitor — penggabung: hasil parser (import.js) → input calc.js → hasil per item.
 * Fungsi murni. Aturan yang sudah DIPUTUSKAN Bagas (16 Nov 2026):
 *   - Demand = SUM qty, jendela 6 bulan BERGULIR (tanggal stok mundur 6 bulan)
 *   - Stock  = semua SLOC (gudang utama 1000/3000 + DMS 5101 + DPS Langkat 5102)
 *   - Oli Bahan "-B" = barang yang sama dengan oli Part (dipindah ke Bahan agar bisa dikeluarkan pecahan liter
 *     untuk paket bundling) → demand & stok DIGABUNG ke kode Part, dihitung & dipesan sebagai PART (ATPM).
 */
(function (root) {
  'use strict';
  const isNode = typeof module !== 'undefined' && module.exports;
  const calc = isNode ? require('./calc.js') : root.LAEPart.calc;
  const importer = isNode ? require('./import.js') : root.LAEPart.importer;

  const SLOC_LABEL = { '1000': 'gudang', '3000': 'gudang', '5101': 'dms', '5102': 'dps' };

  /** Override hari (Stock Level) per pola kode, sesuai Excel: oli KIXX 15+15+15. Bisa di-override dari master. */
  const DEFAULT_DAYS = [[/^SAN-OIL-KIX-/, [15, 15, 15]]];
  const daysFor = (code) => (DEFAULT_DAYS.find(([re]) => re.test(code)) || [])[1];

  /** Kode bahan → kode part tujuan gabung. Bisa di-override dari tab `items` (kolom merge_into). */
  const DEFAULT_MERGE = {
    'D90000-90085-001-B': 'D90000-90085-001', // Oli Synt 0W-20 SN
    'D90000-90055-001-B': 'D90000-90055-002', // Oli Semi Synt 10W-40 API SN
    'D90000-90065-001-B': 'D90000-90065-002', // Oli Semi Synt 5W-30 API SN
  };

  /** 'YYYY-MM-DD' mundur n bulan → awal jendela bergulir (sama seperti tarikan ADMS: 16 Nov → 16 Mei). */
  function rollingFrom(toISO, months = 6) {
    const [y, m, d] = toISO.split('-').map(Number);
    const last = new Date(Date.UTC(y, m - 1 - months + 1, 0)).getUTCDate(); // hari terakhir bulan tujuan
    const dt = new Date(Date.UTC(y, m - 1 - months, Math.min(d, last)));
    return dt.toISOString().slice(0, 10);
  }

  /**
   * parsed: { partSales, bahanIssue, stockPart, stockBahan } — masing-masing hasil importer.parseReport()
   * opts:   { asOf?, months?, merge?, master? (map item_code → {days, moq, active}) , params?,
   *           onOrder? (map item_code → qty PO yang belum datang; kode sudah digabung, lihat po.js) }
   */
  function buildResults(parsed, opts = {}) {
    const months = opts.months || 6;
    const asOf = opts.asOf || parsed.stockPart?.snapshot_date || parsed.stockBahan?.snapshot_date
      || parsed.partSales?.range?.to || parsed.bahanIssue?.range?.to;
    if (!asOf) throw new Error('Tanggal acuan (asOf) tidak diketahui');
    const from = opts.from || rollingFrom(asOf, months);
    const merge = { ...DEFAULT_MERGE, ...(opts.merge || {}) };
    const master = opts.master || {};

    const items = {}; // item_code → agregat
    const get = (code, category, name) => {
      const k = merge[code] || code;
      const cat = merge[code] ? 'PART' : category;
      const it = (items[k] ||= { item_code: k, item_name: '', category: cat, demand: 0, stock: 0,
        stock_gudang: 0, stock_dms: 0, stock_dps: 0, demand_part: 0, demand_bahan: 0, unit_price: null, merged_from: [] });
      if (!it.item_name || (k === code && name)) it.item_name = name || it.item_name;
      if (merge[code] && !it.merged_from.includes(code)) it.merged_from.push(code);
      if (cat === 'PART') it.category = 'PART';
      return it;
    };

    // Demand (SUM, jendela bergulir)
    const addTrx = (rep, category) => {
      for (const t of rep?.rows || []) {
        if (t.date < from || t.date > asOf) continue;
        const it = get(t.item_code, category, t.item_name);
        it.demand += t.qty;
        it[category === 'PART' ? 'demand_part' : 'demand_bahan'] += t.qty;
      }
    };
    addTrx(parsed.partSales, 'PART');
    addTrx(parsed.bahanIssue, 'BAHAN');

    // Stok (semua SLOC) + rincian lokasi
    const addStock = (rep, category) => {
      for (const s of rep?.rows || []) {
        const it = get(s.item_code, category, s.item_name);
        it.stock += s.stock;
        const loc = SLOC_LABEL[s.warehouse];
        if (loc) it['stock_' + loc] += s.stock;
        // harga: pakai harga item part tujuan bila ada
        if (it.unit_price == null || s.item_code === it.item_code) it.unit_price = s.unit_price;
      }
    };
    addStock(parsed.stockPart, 'PART');
    addStock(parsed.stockBahan, 'BAHAN');

    const r12 = (x) => Number.parseFloat(Number(x).toPrecision(12));
    const out = [];
    for (const it of Object.values(items)) {
      const m = master[it.item_code] || {};
      if (m.active === false) continue;
      for (const f of ['demand', 'stock', 'stock_gudang', 'stock_dms', 'stock_dps', 'demand_part', 'demand_bahan']) it[f] = r12(it[f]);
      const res = calc.calcItem({ ...it, on_order: (opts.onOrder || {})[it.item_code] || 0,
        days: m.days || (it.category === 'BAHAN' ? daysFor(it.item_code) : undefined), moq: m.moq }, opts.params);
      res.soq_value = res.unit_price != null ? Math.round(res.soq_order * res.unit_price) : null;
      out.push(res);
    }
    out.sort((a, b) => (b.soq_value || 0) - (a.soq_value || 0) || a.item_code.localeCompare(b.item_code));
    return { asOf, from, to: asOf, count: out.length, items: out };
  }

  /**
   * Hitung ulang SOQ & status dari hasil yang sudah ada (mis. dibaca dari server) dengan On Order terbaru,
   * tanpa perlu file ADMS. Demand & stok tidak berubah; hanya on_order, soq, soq_order, status, soq_value.
   */
  function applyOnOrder(items, onOrder = {}, opts = {}) {
    return items.map((it) => {
      const input = { ...it, on_order: onOrder[it.item_code] || 0 };
      if (!input.days && it.category === 'BAHAN' && daysFor(it.item_code)) input.days = daysFor(it.item_code);
      const res = calc.calcItem(input, opts.params);
      res.soq_value = res.unit_price != null ? Math.round(res.soq_order * res.unit_price) : null;
      return res;
    });
  }

  const api = { buildResults, applyOnOrder, rollingFrom, DEFAULT_MERGE, DEFAULT_DAYS, SLOC_LABEL };
  if (isNode) module.exports = api;
  else root.LAEPart = Object.assign(root.LAEPart || {}, { build: api });
})(typeof window !== 'undefined' ? window : globalThis);
