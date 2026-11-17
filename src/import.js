/**
 * Daya-LAE Part Monitor — parser laporan MENTAH ADMS (Crystal Reports .xls), tanpa dirapikan manual.
 * Fungsi murni: input = array baris (array of arrays) dari SheetJS
 *   XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })   // dengan XLSX.read(..., { cellDates: true })
 *
 * Laporan yang didukung:
 *   PART_SALES   — "Laporan Penjualan Part Gross"   (RptPenjualanPart)
 *   BAHAN_ISSUE  — "Laporan Pengeluaran Bahan Gross" (RptPengeluaranBahan, issue type MIPK)
 *   STOCK_STATUS — "STOCK STATUS"                    (RptStockStatus, Part & Bahan)
 *
 * Pola laporan transaksi: 1 baris header dokumen (tanggal, no dokumen, customer) diikuti baris-baris item.
 * Parser membawa (carry) info header ke setiap baris item. Kolom dicari dari label header, bukan posisi tetap.
 */
(function (root) {
  'use strict';

  const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim().toUpperCase();
  const code = (s) => String(s ?? '').trim().toUpperCase();
  const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

  /** Terima Date, 'dd.mm.yyyy', 'yyyy-mm-dd', atau serial Excel → 'YYYY-MM-DD' */
  function toISODate(v) {
    if (v instanceof Date && !isNaN(v)) {
      // SheetJS cellDates memberi tanggal lokal tengah malam; ambil komponen lokal agar tidak geser hari
      const y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
      return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }
    if (typeof v === 'string') {
      let m = v.trim().match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
      if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
      m = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    }
    if (isNum(v) && v > 20000 && v < 80000) { // serial Excel
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.toISOString().slice(0, 10);
    }
    return null;
  }

  function rowText(rows, n = 40) {
    return rows.slice(0, n).map((r) => (r || []).map(norm).join(' ')).join(' ');
  }

  function detectReport(rows) {
    const t = rowText(rows);
    if (t.includes('PENJUALAN PART')) return 'PART_SALES';
    if (t.includes('PENGELUARAN BAHAN')) return 'BAHAN_ISSUE';
    if (t.includes('STOCK STATUS')) return 'STOCK_STATUS';
    return null;
  }

  /** Cari baris header yang memuat semua label; kembalikan {row, col:{label:index}} */
  function findHeader(rows, labels, maxRow = 60) {
    for (let i = 0; i < Math.min(rows.length, maxRow); i++) {
      const r = rows[i] || [];
      const col = {};
      r.forEach((v, j) => { const k = norm(v); if (k && !(k in col)) col[k] = j; }); // petakan SEMUA label di baris header
      if (labels.every((l) => l in col)) return { row: i, col };
    }
    return null;
  }

  /** Ambil nilai di kolom j; jika kosong (sel merged di Crystal), geser ke kanan maks 2 kolom */
  function near(r, j, pred = (v) => v != null) {
    for (let k = j; k <= j + 2; k++) if (pred(r[k])) return r[k];
    return null;
  }

  function periodRange(rows) {
    const t = rowText(rows, 20);
    const m = t.match(/(\d{2}[./-]\d{2}[./-]\d{4}|\d{4}-\d{2}-\d{2}).{0,10}(S\/D|S\.D).{0,10}?(\d{2}[./-]\d{2}[./-]\d{4}|\d{4}-\d{2}-\d{2})/);
    if (m) return { from: toISODate(m[1]), to: toISODate(m[3]) };
    // Laporan Bahan: tanggal berupa sel Date terpisah
    for (const r of rows.slice(0, 20)) {
      const ds = (r || []).filter((v) => v instanceof Date);
      if (ds.length === 2 && (r || []).some((v) => /S\.?D/i.test(String(v)))) return { from: toISODate(ds[0]), to: toISODate(ds[1]) };
    }
    return { from: null, to: null };
  }

  // ---------- PART_SALES ----------
  function parsePartSales(rows) {
    const h = findHeader(rows, ['NO.PART', 'DESKRIPSI', 'QTY']);
    const hb = findHeader(rows, ['NO.BILLING', 'TGL.BILLING', 'TIPE BILLING']);
    if (!h || !hb) throw new Error('Format Laporan Penjualan Part tidak dikenali (header NO.PART/QTY tidak ditemukan)');
    const c = h.col, cb = hb.col;
    const out = [], errors = [];
    let doc = null;
    for (let i = Math.max(h.row, hb.row) + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const first = r[cb['NO.BILLING']];
      if (typeof first === 'string' && /^\d{4}[A-Z]{3}\d+/.test(first.trim())) {
        doc = { doc_no: first.trim(), date: toISODate(r[cb['TGL.BILLING']]), billing_type: r[cb['TIPE BILLING']] ?? null,
          ref_no: r[cb['NO.REFERENSI']] ?? null,
          // posisi CUSTOMER bergeser 1 kolom tergantung tipe billing (ZPF1 vs ZSFR)
          customer: r[cb['CUSTOMER']] ?? r[cb['CUSTOMER'] - 1] ?? null };
        if (!doc.date) errors.push({ row: i + 1, msg: 'Tanggal billing tidak terbaca', value: r[cb['TGL.BILLING']] });
        continue;
      }
      const part = r[c['NO.PART']];
      if (typeof part === 'string' && part.trim() && isNum(r[c['NO.']] ?? r[c['NO.PART'] - 1])) {
        const qty = near(r, c['QTY'], isNum);
        if (!doc) { errors.push({ row: i + 1, msg: 'Item tanpa header billing' }); continue; }
        if (!isNum(qty)) { errors.push({ row: i + 1, msg: 'Qty bukan angka', value: r[c['QTY']] }); continue; }
        out.push({ source: 'PART_SALES', date: doc.date, period: doc.date && doc.date.slice(0, 7), item_code: code(part),
          item_name: String(r[c['DESKRIPSI']] ?? '').trim(), qty, uom: null, doc_no: doc.doc_no, trx_type: doc.billing_type,
          ref_no: doc.ref_no, customer: doc.customer });
      }
    }
    return { report: 'PART_SALES', range: periodRange(rows), rows: out, errors };
  }

  // ---------- BAHAN_ISSUE ----------
  function parseBahanIssue(rows) {
    const h = findHeader(rows, ['NO', 'NOMOR - NAMA BAHAN', 'QTY.', 'SATUAN']);
    const hd = findHeader(rows, ['TANGGAL', 'MOVEMENT TYPE', 'NO. DOKUMEN']);
    if (!h || !hd) throw new Error('Format Laporan Pengeluaran Bahan tidak dikenali');
    const c = h.col, cd = hd.col;
    const out = [], errors = [];
    let doc = null;
    for (let i = Math.max(h.row, hd.row) + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const d = r[cd['TANGGAL']];
      const mv = r[cd['MOVEMENT TYPE']];
      if (toISODate(d) && mv != null) {
        doc = { date: toISODate(d), movement: String(mv).trim(), doc_no: r[cd['NO. DOKUMEN']] ?? null,
          ref_no: r[cd['NO. REFERENSI']] ?? null, customer: r[cd['KODE - CUSTOMER NAME']] ?? null };
        continue;
      }
      const nb = r[c['NOMOR - NAMA BAHAN']];
      if (isNum(r[c['NO']]) && typeof nb === 'string' && nb.includes(' - ')) {
        const k = nb.indexOf(' - ');
        const qty = near(r, c['QTY.'], isNum);
        if (!doc) { errors.push({ row: i + 1, msg: 'Item tanpa header dokumen' }); continue; }
        if (!isNum(qty)) { errors.push({ row: i + 1, msg: 'Qty bukan angka', value: r[c['QTY.']] }); continue; }
        // movement 262 = pembatalan/retur issue → qty negatif
        const sign = doc.movement === '262' ? -1 : 1;
        out.push({ source: 'BAHAN_ISSUE', date: doc.date, period: doc.date.slice(0, 7), item_code: code(nb.slice(0, k)),
          item_name: nb.slice(k + 3).trim(), qty: sign * qty, uom: near(r, c['SATUAN'], (v) => typeof v === 'string'),
          doc_no: doc.doc_no, trx_type: doc.movement, ref_no: doc.ref_no, customer: doc.customer });
      }
    }
    return { report: 'BAHAN_ISSUE', range: periodRange(rows), rows: out, errors };
  }

  // ---------- STOCK_STATUS ----------
  function parseStockStatus(rows) {
    const h = findHeader(rows, ['NO', 'MATERIAL NO', 'MATERIAL DESCRIPTION', 'SLOC', 'UNR. QTY.']);
    if (!h) throw new Error('Format Stock Status tidak dikenali');
    const c = h.col, out = [], errors = [];
    const get = (r, k) => (k in c ? r[c[k]] : null);
    for (let i = h.row + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const mat = get(r, 'MATERIAL NO');
      if (!isNum(get(r, 'NO')) || typeof mat !== 'string' || !mat.trim()) continue;
      const qty = get(r, 'UNR. QTY.');
      if (!isNum(qty)) { errors.push({ row: i + 1, msg: 'UNR. QTY bukan angka', value: qty }); continue; }
      out.push({ item_code: code(mat), item_name: String(get(r, 'MATERIAL DESCRIPTION') ?? '').trim(),
        plant: String(get(r, 'PLANT') ?? ''), warehouse: String(get(r, 'SLOC') ?? ''), bin: get(r, 'STORAGE BIN'),
        unit_price: get(r, 'UNIT PRICE'), stock: qty, qty_bin: get(r, 'QTY. BIN.'), so_qty: get(r, 'SO QTY.'),
        vendor_qty: get(r, 'VENDOR QTY.'), product_class: get(r, 'PRODUCT CLASS') });
    }
    const tgl = (() => { for (const r of rows.slice(0, 10)) { const j = (r || []).findIndex((v) => norm(v) === 'TANGGAL' || norm(v) === 'TANGGAL :'); if (j >= 0) for (let k = j + 1; k < j + 4; k++) { const d = toISODate(r[k]); if (d) return d; } } return null; })();
    return { report: 'STOCK_STATUS', snapshot_date: tgl, rows: out, errors };
  }

  function parseReport(rows) {
    const t = detectReport(rows);
    if (t === 'PART_SALES') return parsePartSales(rows);
    if (t === 'BAHAN_ISSUE') return parseBahanIssue(rows);
    if (t === 'STOCK_STATUS') return parseStockStatus(rows);
    throw new Error('Jenis laporan tidak dikenali. Gunakan export ADMS: Penjualan Part, Pengeluaran Bahan, atau Stock Status.');
  }

  /** Jumlahkan stok per item. warehouses: null = semua gudang; ['3000'] = hanya SLOC 3000 */
  function stockByItem(stockRows, warehouses = null) {
    const out = {};
    for (const s of stockRows) {
      if (warehouses && !warehouses.includes(s.warehouse)) continue;
      out[s.item_code] = (out[s.item_code] || 0) + s.stock;
    }
    for (const k in out) out[k] = Number.parseFloat(out[k].toPrecision(12));
    return out;
  }

  /** Demand per item dalam rentang tanggal [from, to] (inklusif, 'YYYY-MM-DD'). method SUM | COUNT */
  function demandByItem(trxRows, { from, to, method = 'SUM' } = {}) {
    const out = {};
    for (const t of trxRows) {
      if (from && t.date < from) continue;
      if (to && t.date > to) continue;
      out[t.item_code] = (out[t.item_code] || 0) + (method === 'COUNT' ? 1 : t.qty);
    }
    for (const k in out) out[k] = Number.parseFloat(out[k].toPrecision(12));
    return out;
  }

  /** Demand per item per bulan (untuk grafik & pengecekan dengan pivot Excel) */
  function demandMonthly(trxRows, method = 'SUM') {
    const out = {};
    for (const t of trxRows) {
      const m = (out[t.item_code] ||= {});
      m[t.period] = (m[t.period] || 0) + (method === 'COUNT' ? 1 : t.qty);
    }
    return out;
  }

  const api = { toISODate, detectReport, parseReport, parsePartSales, parseBahanIssue, parseStockStatus, stockByItem, demandByItem, demandMonthly };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LAEPart = Object.assign(root.LAEPart || {}, { importer: api });
})(typeof window !== 'undefined' ? window : globalThis);
