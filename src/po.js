/**
 * Daya-LAE Part Monitor — PO outstanding (fungsi murni, tanpa DOM/fetch).
 *
 * Sumber data: screenshot layar ADMS "Maintain Purchase Order" (daftar) dan "Display Purchase Order" (detail),
 * dibaca OCR menjadi `lines` = array baris, tiap baris = array sel { text, box:{x,y,width,height}, ink? }.
 * (`ink` = jumlah piksel abu-abu di dalam sel, diisi UI; dipakai untuk membaca centang kolom Dlv.)
 *
 * Aturan bisnis (keputusan Bagas, 17 Nov 2026):
 *   - Pengurang SOQ ("On Order") = baris PO yang BELUM datang dari:
 *       ZPOR (order reguler ATPM)  : status Open + ada SAP Doc No., atau Partial (baris yang belum tercentang Dlv)
 *       PART (vendor lokal)        : status Open / Partial
 *       ZCON (PO Bahan)            : status Open / Partial, tanggal PO dalam jendela (default 6 bulan terakhir)
 *   - ZPOR Open TANPA SAP Doc No. = PO gagal → diabaikan.
 *   - Completed = sudah di-GR, stok ADMS sudah naik → ditutup.
 *   - ZPOL (warranty / pesanan khusus, terikat PKB) TIDAK mengurangi SOQ; dipantau terpisah agar tidak jadi dead stock.
 *   - Partial: per baris, datang penuh atau belum sama sekali (centang Dlv).
 */
(function (root) {
  'use strict';

  const INK_RATIO = 1.4; // kotak tercentang bila tintanya ≥ 1,4× kotak paling kosong di tabel itu
  const STATUS = { OPEN: 'Open', PARTIAL: 'Partial', COMPLETED: 'Completed' };

  const up = (s) => String(s ?? '').toUpperCase();
  /** Rapikan salah baca OCR yang umum pada nomor dokumen: "3401SPA,26005855", "3401 SBC25000019". */
  /* Juga: huruf B kadang terbaca sebagai aksara "日", dan seri dokumen tersisipi koma ("3401SP,A,26005855"). */
  const tidy = (s) => up(s).replace(/日/g, 'B')
    .replace(/(\d{4})\s?([A-Z][A-Z,.\s]{1,6})(\d{8})/g, (m, a, b, c) => { const t = b.replace(/[^A-Z]/g, ''); return t.length === 3 ? a + t + c : m; })
    .replace(/ZP0/g, 'ZPO').replace(/ZC0N/g, 'ZCON');
  const rowText = (row) => tidy((row || []).map((c) => c.text || '').join(' '));
  const center = (c) => c.box.x + c.box.width / 2;

  function findDoc(text, wantPkb) {
    const re = /(\d{4})([A-Z]{3})(\d{8})/g;
    let m;
    while ((m = re.exec(text))) if ((m[2] === 'SPA') === !!wantPkb) return { no: m[0], series: m[2], index: m.index, end: m.index + m[0].length };
    return null;
  }
  function findDate(text) {
    const m = /(\d{2})\s?\.\s?(\d{2})\s?\.\s?(\d{4})/.exec(text);
    return m ? { iso: `${m[3]}-${m[2]}-${m[1]}`, end: m.index + m[0].length } : null;
  }
  const findType = (text) => (/(ZPOR|ZPOL|ZCON|PART(?!IAL))/.exec(text) || [])[1] || '';
  function findStatus(text) {
    if (/C\w{0,3}MP\w{0,2}ETED/.test(text)) return STATUS.COMPLETED;
    if (/PARTIAL|PART[1I]AL/.test(text)) return STATUS.PARTIAL;
    if (/\bOPEN|OPEN\b|0PEN/.test(text)) return STATUS.OPEN;
    return '';
  }
  const kindOf = (type, series) => (type === 'ZCON' || series === 'SBC' ? 'BAHAN' : 'PART');

  /** "5" → 5 · "720,000" → 720 · "1.234,5" → 1234.5 (ADMS: titik = ribuan, koma = desimal) */
  function parseQty(s) {
    const t = String(s ?? '').replace(/[^\d.,]/g, '');
    if (!t || !/\d/.test(t)) return null;
    const n = Number(t.replace(/\./g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  }

  // ---------- Koreksi kode item terhadap daftar item yang dikenal ----------
  const canon = (code) => up(code).replace(/[O0DQ]/g, '0').replace(/[I1L|]/g, '1').replace(/[S5]/g, '5').replace(/[B8]/g, '8').replace(/[Z2]/g, '2').replace(/[G6]/g, '6');
  function makeMatcher(knownCodes) {
    const known = new Set((knownCodes || []).map(up));
    const byCanon = {};
    for (const k of known) (byCanon[canon(k)] ||= []).push(k);
    /** → { code, known, corrected } — hanya mengoreksi bila tepat SATU kode dikenal yang mirip */
    return (raw) => {
      const code = up(raw).replace(/\s+/g, '').replace(/[^A-Z0-9-]/g, '');
      if (known.has(code)) return { code, known: true, corrected: false };
      const cand = byCanon[canon(code)] || [];
      if (cand.length === 1) return { code: cand[0], known: true, corrected: true, read: code };
      return { code, known: false, corrected: false };
    };
  }

  function detectScreen(lines) {
    const t = (lines || []).slice(0, 30).map(rowText).join(' ');
    if (/L[I1]ST OF PURCHASE ORDER/.test(t) || /MAINTAIN PURCHASE ORDER/.test(t)) return 'LIST';
    if (/DISPLAY PURCHASE ORDER/.test(t) || /MATERIAL N[O0]/.test(t) || /ITE\s?M DETAILS/.test(t)) return 'DETAIL';
    return null;
  }

  // ---------- Layar daftar PO ----------
  function parseList(lines) {
    const head = (lines || []).find((r) => r.some((c) => /SAP\s*DOC/i.test(c.text || '')));
    const sapCell = head && head.find((c) => /SAP\s*DOC/i.test(c.text || ''));
    const orders = [];
    for (const row of lines || []) {
      const t = rowText(row);
      const po = findDoc(t, false);
      if (!po) continue;
      const rest = t.slice(po.end);
      const date = findDate(rest);
      const tail = date ? rest.slice(date.end) : rest;
      const type = findType(tail) || (po.series === 'SBC' ? 'ZCON' : '');
      const pkb = findDoc(t, true);
      let sap = '';
      for (const c of row) {
        const v = String(c.text || '').trim();
        if (!/^\d{10}$/.test(v)) continue;
        if (sapCell ? c.box.x >= sapCell.box.x - 80 : c === row[row.length - 1]) sap = v;
      }
      orders.push({ po_no: po.no, po_date: date ? date.iso : '', po_type: type, status: findStatus(tail), sap_doc: sap,
        pkb_no: pkb ? pkb.no : '', kind: kindOf(type, po.series) });
    }
    return { screen: 'LIST', orders };
  }

  // ---------- Layar detail PO ----------
  /**
   * Kotak centang dibaca dari "tinta" (piksel abu-abu) di dalam sel, dibandingkan dengan kotak paling kosong di tabel yang sama:
   * kotak tercentang ≈ 1,6–1,9× kotak kosong, apa pun skala/penghalusan gambarnya. Tanpa data tinta → pakai glif OCR.
   */
  function ticked(cell, refInk) {
    if (!cell) return false;
    if (typeof cell.ink === 'number' && refInk > 0) return cell.ink >= refInk * INK_RATIO;
    return /[区☑√✓Vv]/.test(cell.text || '');
  }

  function parseDetail(lines, opts = {}) {
    const match = opts.match || makeMatcher(opts.known);
    const rows = lines || [];
    const hi = rows.findIndex((r) => r.some((c) => /MATERIAL\s*N[O0]/i.test(c.text || '')) && r.some((c) => /^\W*Q[T1I]Y/i.test(c.text || '')));
    const headRows = hi >= 0 ? rows.slice(0, hi) : rows;
    const headText = headRows.map(rowText).join(' \n ');
    const after = (re) => { // teks sel label + sel-sel sesudahnya di baris yang sama
      for (const r of headRows) { const i = r.findIndex((c) => re.test(c.text || '')); if (i >= 0) return tidy(r.slice(i, i + 3).map((c) => c.text || '').join(' ')); }
      return '';
    };
    const po = findDoc(headText, false);
    const pkb = findDoc(headText, true);
    const type = findType(after(/ORDER\s*TYPE/i)) || findType(headText) || (po && po.series === 'SBC' ? 'ZCON' : '');
    const date = findDate(after(/ORDER\s*DATE/i)) || findDate(headText);
    const chassis = (/\b([A-HJ-NPR-Z0-9]{17})\b/.exec(after(/CHASSIS/i)) || [])[1] || '';
    const order = { po_no: po ? po.no : '', po_date: date ? date.iso : '', po_type: type, status: findStatus(after(/^\W*STATUS/i)),
      sap_doc: '', pkb_no: pkb ? pkb.no : '', chassis_no: chassis, kind: kindOf(type, po ? po.series : ''), lines: [], has_detail: true };
    if (hi < 0) return { screen: 'DETAIL', order, problems: ['Tabel item tidak ditemukan di gambar.'] };

    const head = rows[hi];
    const matHead = head.find((c) => /MATERIAL\s*N[O0]/i.test(c.text || ''));
    const qtyHead = head.find((c) => /^\W*Q[T1I]Y/i.test(c.text || ''));
    const delHead = head.find((c) => /^\W*DE[L1I]\b|^\W*DE[L1I]\./i.test(c.text || ''));
    const dlvHead = delHead ? head[head.indexOf(delHead) + 1] : null; // kolom sesudah "Del." (terbaca "Dlv.", "Dl1", "DI1")
    const problems = [];
    const body = [];
    for (const row of rows.slice(hi + 1)) { if (/PRINT\s*PREVIEW|PO\s*HISTORY/i.test(rowText(row))) break; body.push(row); }
    const boxOf = (row, h) => { // sel di baris ini yang paling dekat dengan judul kolom h
      if (!h) return null;
      let best = null;
      for (const c of row) { const d = Math.abs(center(c) - center(h)); if (d <= Math.max(h.box.width, 40) && (!best || d < Math.abs(center(best) - center(h)))) best = c; }
      return best;
    };
    const inks = [];
    for (const row of body) for (const h of [delHead, dlvHead]) { const c = boxOf(row, h); if (c && typeof c.ink === 'number' && c.ink > 0) inks.push(c.ink); }
    const refInk = inks.length ? Math.min(...inks) : 0;
    for (const row of body) {
      // sel kode = sel "mirip kode" yang paling dekat dengan judul kolom Material No., di kiri kolom Qty
      let codeCell = null;
      for (const c of row) {
        const v = up(c.text).replace(/\s+/g, '');
        if (v.length < 5 || !/\d/.test(v) || !/^[A-Z0-9]+([-.][A-Z0-9]+)*$/.test(v) || /^\d+$/.test(v)) continue;
        if (qtyHead && c.box.x >= qtyHead.box.x) continue;
        if (!codeCell || Math.abs(center(c) - center(matHead)) < Math.abs(center(codeCell) - center(matHead))) codeCell = c;
      }
      if (!codeCell) continue;
      const i = row.indexOf(codeCell);
      const rest = row.slice(i + 1);
      const qtyCell = rest.find((c) => /^\D{0,2}\d[\d.,]*\D{0,1}$/.test(String(c.text || '').trim()));
      const descCell = rest.find((c) => c !== qtyCell && String(c.text || '').trim().length > 1);
      const qty = qtyCell ? parseQty(qtyCell.text) : null;
      const m = match(codeCell.text);
      const delCell = boxOf(row, delHead), dlvCell = boxOf(row, dlvHead);
      // kolom Del. tercentang = baris dihapus dari PO → tidak dihitung
      if (delCell && delCell !== dlvCell && typeof delCell.ink === 'number' && ticked(delCell, refInk)) continue;
      const delivered = dlvCell && dlvCell !== delCell ? ticked(dlvCell, refInk) : false;
      if (qty == null) problems.push(`Qty ${m.code} tidak terbaca.`);
      order.lines.push({ item_code: m.code, item_name: descCell && rest.indexOf(descCell) < (qtyCell ? rest.indexOf(qtyCell) : 99) ? String(descCell.text).trim() : '',
        qty: qty == null ? 0 : qty, delivered, known: m.known, corrected: m.corrected ? m.read : '' });
    }
    if (!order.po_no) problems.push('No PO tidak terbaca.');
    if (!order.lines.length) problems.push('Tidak ada baris item yang terbaca.');
    return { screen: 'DETAIL', order, problems };
  }

  // ---------- Penyimpanan PO (objek po_no → order) ----------
  /** Gabungkan hasil baca ke simpanan. today = 'YYYY-MM-DD' (tanggal tangkapan). Mengembalikan ringkasan perubahan. */
  function mergeParsed(store, parsed, today) {
    const sum = { added: 0, updated: 0, closed: 0 };
    const upsert = (o, isDetail) => {
      if (!o.po_no) return;
      const cur = store[o.po_no];
      const next = cur ? { ...cur } : { po_no: o.po_no, po_date: '', po_type: '', status: '', sap_doc: '', pkb_no: '', chassis_no: '', kind: o.kind, lines: [], has_detail: false };
      for (const f of ['po_date', 'po_type', 'status', 'sap_doc', 'pkb_no', 'chassis_no']) if (o[f]) next[f] = o[f];
      next.kind = kindOf(next.po_type, o.po_no.slice(4, 7));
      if (isDetail) {
        const old = Object.fromEntries((cur ? cur.lines : []).map((l) => [l.item_code, l]));
        next.lines = o.lines.map((l) => ({ ...l, delivered_at: (old[l.item_code] || {}).delivered_at || '', sold: (old[l.item_code] || {}).sold || '' }));
        next.has_detail = true;
      }
      for (const l of next.lines) {
        if (next.status === STATUS.COMPLETED) l.delivered = true;
        if (l.delivered && !l.delivered_at) l.delivered_at = today;
      }
      next.updated_at = today;
      store[o.po_no] = next;
      cur ? sum.updated++ : sum.added++;
    };
    if (parsed.screen === 'LIST') parsed.orders.forEach((o) => upsert(o, false));
    else if (parsed.screen === 'DETAIL') upsert(parsed.order, true);
    // PO biasa yang sudah Completed tidak perlu disimpan lagi (stok ADMS sudah naik); ZPOL tetap dipantau.
    for (const k of Object.keys(store)) if (store[k].status === STATUS.COMPLETED && store[k].po_type !== 'ZPOL') { delete store[k]; sum.closed++; }
    return sum;
  }

  /**
   * Peran sebuah PO dalam hitungan:
   *   COUNT (mengurangi SOQ) · NEED_DETAIL (akan dihitung, tapi baris item belum ada) · FAILED · OLD · CLOSED · SPECIAL (ZPOL)
   */
  function classify(order, opts = {}) {
    if (order.po_type === 'ZPOL') return 'SPECIAL';
    if (order.status === STATUS.COMPLETED) return 'CLOSED';
    if (order.kind === 'BAHAN') { if (opts.bahanFrom && order.po_date && order.po_date < opts.bahanFrom) return 'OLD'; }
    else if (order.po_type !== 'PART' && order.status !== STATUS.PARTIAL && !order.sap_doc) return 'FAILED';
    return order.has_detail && order.lines.length ? 'COUNT' : 'NEED_DETAIL';
  }

  /** → { map: {item_code: qty belum datang}, counted: [po_no], needDetail: [po_no] } */
  function onOrderByItem(store, opts = {}) {
    const merge = opts.merge || {};
    const map = {}, counted = [], needDetail = [];
    for (const o of Object.values(store || {})) {
      const c = classify(o, opts);
      if (c === 'NEED_DETAIL') needDetail.push(o.po_no);
      if (c !== 'COUNT') continue;
      counted.push(o.po_no);
      for (const l of o.lines) {
        if (l.delivered || !(l.qty > 0) || !l.item_code) continue;
        const k = merge[l.item_code] || l.item_code;
        map[k] = Number.parseFloat(((map[k] || 0) + l.qty).toPrecision(12));
      }
    }
    return { map, counted, needDetail };
  }

  /** Tandai baris ZPOL yang part-nya sudah terjual pada PKB yang sama (no referensi penjualan = no PKB). */
  function markSold(store, salesRows) {
    const sold = {};
    for (const t of salesRows || []) if (t.ref_no) { const k = String(t.ref_no).trim() + '|' + t.item_code; if (!sold[k] || t.date < sold[k]) sold[k] = t.date; }
    let n = 0;
    for (const o of Object.values(store || {})) {
      if (o.po_type !== 'ZPOL' || !o.pkb_no) continue;
      for (const l of o.lines) { const d = sold[o.pkb_no + '|' + l.item_code]; if (d && l.sold !== d) { l.sold = d; n++; } }
    }
    return n;
  }

  const daysBetween = (a, b) => Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);

  /**
   * Baris pantauan pesanan khusus (ZPOL). stage:
   *   ORDERED (belum datang) · ARRIVED (sudah datang, unit belum dikerjakan) · DONE (sudah terpasang/terjual) · NEED_DETAIL
   *   risk = ARRIVED lebih dari limitDays hari → berpotensi dead stock.
   */
  function specialRows(store, opts = {}) {
    const today = opts.today, limit = opts.limitDays == null ? 14 : opts.limitDays;
    const out = [];
    for (const o of Object.values(store || {})) {
      if (o.po_type !== 'ZPOL') continue;
      const base = { po_no: o.po_no, po_date: o.po_date, pkb_no: o.pkb_no, chassis_no: o.chassis_no };
      if (!o.has_detail || !o.lines.length) { out.push({ ...base, item_code: '', item_name: '', qty: null, stage: 'NEED_DETAIL', age: null, risk: false }); continue; }
      for (const l of o.lines) {
        let stage = 'ORDERED', since = o.po_date;
        if (l.sold) { stage = 'DONE'; since = l.sold; }
        else if (l.delivered) { stage = 'ARRIVED'; since = l.delivered_at || o.updated_at || o.po_date; }
        const age = today && since ? Math.max(0, daysBetween(since, today)) : null;
        out.push({ ...base, item_code: l.item_code, item_name: l.item_name, qty: l.qty, stage, age, risk: stage === 'ARRIVED' && age != null && age > limit });
      }
    }
    const rank = { ARRIVED: 0, ORDERED: 1, NEED_DETAIL: 2, DONE: 3 };
    return out.sort((a, b) => (b.risk - a.risk) || rank[a.stage] - rank[b.stage] || (b.age || 0) - (a.age || 0));
  }

  // ---------- Bentuk baris untuk tab Sheets `on_order` (satu baris per item PO) ----------
  const ROW_FIELDS = ['po_no', 'po_date', 'po_type', 'status', 'sap_doc', 'pkb_no', 'chassis_no', 'item_code', 'item_name', 'qty', 'delivered', 'delivered_at', 'sold', 'updated_at'];
  function toRows(store) {
    const rows = [];
    for (const o of Object.values(store || {})) {
      const head = { po_no: o.po_no, po_date: o.po_date, po_type: o.po_type, status: o.status, sap_doc: o.sap_doc, pkb_no: o.pkb_no, chassis_no: o.chassis_no, updated_at: o.updated_at || '' };
      if (!o.has_detail || !o.lines.length) rows.push({ ...head, item_code: '', item_name: '', qty: '', delivered: '', delivered_at: '', sold: '' });
      for (const l of o.has_detail ? o.lines : []) rows.push({ ...head, item_code: l.item_code, item_name: l.item_name || '', qty: l.qty, delivered: !!l.delivered, delivered_at: l.delivered_at || '', sold: l.sold || '' });
    }
    return rows;
  }
  function fromRows(rows) {
    const store = {};
    for (const r of rows || []) {
      const no = String(r.po_no || '').trim();
      if (!no) continue;
      const s = (v) => (v == null ? '' : String(v).trim());
      const o = (store[no] ||= { po_no: no, po_date: s(r.po_date), po_type: s(r.po_type), status: s(r.status), sap_doc: s(r.sap_doc), pkb_no: s(r.pkb_no),
        chassis_no: s(r.chassis_no), kind: kindOf(s(r.po_type), no.slice(4, 7)), lines: [], has_detail: false, updated_at: s(r.updated_at) });
      if (s(r.item_code)) {
        o.has_detail = true;
        o.lines.push({ item_code: s(r.item_code).toUpperCase(), item_name: s(r.item_name), qty: Number(r.qty) || 0, delivered: r.delivered === true || /^true$/i.test(s(r.delivered)),
          delivered_at: s(r.delivered_at), sold: s(r.sold), known: true, corrected: '' });
      }
    }
    return store;
  }

  const api = { detectScreen, parseList, parseDetail, parseQty, makeMatcher, mergeParsed, classify, onOrderByItem, markSold, specialRows, toRows, fromRows, ROW_FIELDS, STATUS, INK_RATIO };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LAEPart = Object.assign(root.LAEPart || {}, { po: api });
})(typeof window !== 'undefined' ? window : globalThis);
