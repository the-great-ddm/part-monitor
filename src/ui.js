/**
 * Daya-LAE Part Monitor — UI: upload export ADMS → hitung → kartu ringkasan + tabel + export SOQ.
 * Semua perhitungan ada di calc.js / import.js / build.js; file ini hanya membaca file & menampilkan.
 */
(function (root) {
  'use strict';
  const L = root.LAEPart;
  const doc = root.document;
  const $ = (id) => doc.getElementById(id);
  const STORE_KEY = 'laepart.result.v1';
  const PO_KEY = 'laepart.po.v1';
  const PARAM_KEY = 'laepart.params.v1'; // salinan parameter aktif terakhir (agar tetap jalan tanpa server)
  const PAGE = 300;

  const BULAN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const num = (x, d = 1) => (x == null || !Number.isFinite(x) ? '–' : x.toLocaleString('id-ID', { maximumFractionDigits: d }));
  const rp = (x) => (x == null ? '–' : 'Rp ' + Math.round(x).toLocaleString('id-ID'));
  const rpShort = (x) => {
    if (x >= 1e9) return 'Rp ' + num(x / 1e9, 2) + ' M';
    if (x >= 1e6) return 'Rp ' + num(x / 1e6, 1) + ' jt';
    return rp(x);
  };
  const tgl = (iso) => {
    if (!iso) return '–';
    const [y, m, d] = iso.split('-');
    return `${+d} ${BULAN[+m - 1]} ${y}`;
  };

  // Urutan = prioritas tindakan (PRD §7)
  const STATUS = {
    STOCKOUT: { label: 'Stockout', cls: 's-stockout' },
    ORDER: { label: 'Stock Menipis', cls: 's-order' },
    INCOMING: { label: 'Menunggu Datang', cls: 's-incoming' },
    NORMAL: { label: 'Normal', cls: 's-normal' },
    OVERSTOCK: { label: 'Overstock', cls: 's-overstock' },
    NON_STOCK: { label: 'Non-Stock', cls: 's-nonstock' },
    DEAD: { label: 'Dead', cls: 's-dead' },
    NO_MOVE: { label: 'No Move', cls: 's-nomove' },
  };
  const RANK = Object.fromEntries(Object.keys(STATUS).map((k, i) => [k, i]));

  const SLOTS = {
    partSales: 'Penjualan Part',
    bahanIssue: 'Pengeluaran Bahan',
    stockPart: 'Stock Status Part',
    stockBahan: 'Stock Status Bahan',
  };

  const COLS = [
    { key: 'item_code', label: 'Item', left: true },
    { key: 'category', label: 'Kat.', left: true },
    { key: 'status', label: 'Status', left: true },
    { key: 'stock', label: 'Stock' },
    { key: 'on_order', label: 'On Order' },
    { key: 'demand', label: 'Demand 6 bln' },
    { key: 'mad', label: 'MAD' },
    { key: 'stock_level', label: 'Stock Level' },
    { key: 'mip', label: 'MIP' },
    { key: 'soq', label: 'SOQ' },
    { key: 'soq_order', label: 'Order' },
    { key: 'soq_value', label: 'Nilai Order' },
  ];

  const state = {
    parsed: {}, files: {}, fileErrors: [], result: null, fresh: false,
    filter: { q: '', cat: '', status: '' }, sort: { key: 'priority', dir: 1 }, limit: PAGE,
    po: { store: {}, dirty: false, log: [], busy: false, serverApi: 0 },
    // parameter rumus aktif (menu Parameter, src/paramui.js); log = riwayat perubahan dari server
    params: { active: L.params.clone(L.params.DEFAULT_SETTINGS), log: [], serverApi: 0, fromServer: false },
  };

  // ---------- Baca file ----------
  async function readFile(file) {
    if (!root.XLSX) throw new Error('Pustaka pembaca Excel gagal dimuat. Periksa koneksi internet lalu muat ulang halaman.');
    let wb;
    try {
      wb = root.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    } catch (e) {
      throw new Error('File tidak bisa dibaca sebagai Excel/CSV.');
    }
    for (const name of wb.SheetNames) {
      const rows = root.XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
      if (L.importer.detectReport(rows)) return L.importer.parseReport(rows);
    }
    throw new Error('Jenis laporan tidak dikenali. Gunakan export ADMS: Penjualan Part, Pengeluaran Bahan, atau Stock Status.');
  }

  /** Stock Status Part & Bahan berjudul sama → bedakan dari SLOC gudang utama (1000 = Part, 3000 = Bahan). */
  function slotOf(rep, fileName) {
    if (rep.report === 'PART_SALES') return 'partSales';
    if (rep.report === 'BAHAN_ISSUE') return 'bahanIssue';
    const w = new Set(rep.rows.map((r) => r.warehouse));
    if (w.has('3000') && !w.has('1000')) return 'stockBahan';
    if (w.has('1000') && !w.has('3000')) return 'stockPart';
    return /bahan/i.test(fileName) ? 'stockBahan' : 'stockPart';
  }

  async function addFiles(fileList) {
    state.fileErrors = [];
    $('drop').querySelector('strong').textContent = 'Membaca file…';
    for (const f of Array.from(fileList)) {
      try {
        const rep = await readFile(f);
        const slot = slotOf(rep, f.name);
        state.parsed[slot] = rep;
        state.files[slot] = f.name;
      } catch (e) {
        state.fileErrors.push(`${f.name}: ${e.message}`);
      }
    }
    $('drop').querySelector('strong').textContent = 'Tarik file ke sini, atau klik untuk memilih';
    renderUpload();
  }

  /** Pemeriksaan sebelum hitung: file lengkap, tanggal stok sama, transaksi mencakup jendela 6 bulan. */
  function checks() {
    const p = state.parsed, errors = [...state.fileErrors], warns = [];
    const missing = Object.keys(SLOTS).filter((k) => !p[k]);
    for (const k of Object.keys(SLOTS)) if (p[k] && !p[k].rows.length) errors.push(`${SLOTS[k]}: tidak ada baris data yang terbaca.`);
    const dp = p.stockPart?.snapshot_date, db = p.stockBahan?.snapshot_date;
    if (dp && db && dp !== db) warns.push(`Tanggal Stock Status berbeda: Part ${tgl(dp)}, Bahan ${tgl(db)}. Perhitungan memakai tanggal Part.`);
    const asOf = dp || db;
    if (asOf) {
      const from = L.build.rollingFrom(asOf);
      for (const k of ['partSales', 'bahanIssue']) {
        const r = p[k]?.range;
        if (!r || !r.from || !r.to) continue;
        if (r.from > from) warns.push(`${SLOTS[k]} mulai ${tgl(r.from)}, padahal jendela 6 bulan mulai ${tgl(from)}. Demand & MAD akan terlalu kecil.`);
        if (r.to < asOf) warns.push(`${SLOTS[k]} hanya sampai ${tgl(r.to)}, sedangkan stok per ${tgl(asOf)}.`);
      }
    } else if (p.stockPart || p.stockBahan) {
      warns.push('Tanggal Stock Status tidak terbaca; tanggal acuan diambil dari akhir periode transaksi.');
    }
    return { missing, errors, warns };
  }

  function renderUpload() {
    const p = state.parsed;
    $('slots').innerHTML = Object.entries(SLOTS).map(([k, label]) => {
      const r = p[k];
      if (!r) return `<li class="slot"><span class="mark">–</span><div class="body"><b>${label}</b><small>Belum ada file</small></div></li>`;
      const info = r.report === 'STOCK_STATUS'
        ? `Stok per ${tgl(r.snapshot_date)}`
        : `${tgl(r.range.from)} s/d ${tgl(r.range.to)}`;
      const errs = r.errors.length
        ? `<details><summary>${num(r.errors.length, 0)} baris gagal dibaca</summary><ul>${r.errors.slice(0, 100).map((e) =>
          `<li>Baris ${e.row}: ${esc(e.msg)}${e.value != null ? ` (${esc(e.value)})` : ''}</li>`).join('')}</ul></details>`
        : '';
      return `<li class="slot ${r.errors.length || !r.rows.length ? 'bad' : 'ok'}"><span class="mark">✓</span><div class="body"><b>${label}</b>
        <small>${esc(state.files[k])}</small><small>${info} · ${num(r.rows.length, 0)} baris</small>${errs}</div></li>`;
    }).join('');

    const c = checks();
    const notes = [
      ...c.errors.map((m) => `<li class="err">${esc(m)}</li>`),
      ...c.warns.map((m) => `<li class="warn">${esc(m)}</li>`),
    ];
    const any = Object.keys(p).length > 0;
    if (any && c.missing.length) notes.push(`<li class="info">Belum lengkap: ${c.missing.map((k) => SLOTS[k]).join(', ')}.</li>`);
    $('notes').innerHTML = notes.join('');
    $('btnCalc').disabled = c.missing.length > 0 || Object.keys(SLOTS).some((k) => !p[k].rows.length);
  }

  function resetUpload() {
    state.parsed = {}; state.files = {}; state.fileErrors = [];
    $('fileInput').value = '';
    renderUpload();
  }

  function compute() {
    let r;
    try {
      r = L.build.buildResults(state.parsed, { settings: state.params.active });
    } catch (e) {
      state.fileErrors = ['Perhitungan gagal: ' + e.message];
      renderUpload();
      return;
    }
    // unsaved: hasil upload belum dikirim ke server → jangan ditimpa data server yang lebih lama saat halaman dimuat ulang
    state.result = { asOf: r.asOf, from: r.from, items: r.items, source: 'upload', files: { ...state.files }, unsaved: !!(L.api && L.api.configured()) };
    state.fresh = true;
    // pesanan khusus (ZPOL): part yang sudah terjual pada PKB-nya ditandai selesai
    if (L.po.markSold(state.po.store, state.parsed.partSales && state.parsed.partSales.rows) > 0) { state.po.dirty = true; persistPo(); }
    applyOnOrder();
    try { root.localStorage.setItem(STORE_KEY, JSON.stringify(state.result)); } catch (e) { /* penyimpanan lokal tidak tersedia */ }
    showUpload(false);
    renderDash();
    renderPo();
  }

  // ---------- Dashboard ----------
  const value = (x) => (x.unit_price != null ? x.stock * x.unit_price : 0);

  function summary(items) {
    const s = { total: 0, soqPart: 0, soqBahan: 0, nPart: 0, nBahan: 0, noPrice: 0, by: {} };
    for (const k of Object.keys(STATUS)) s.by[k] = { n: 0, val: 0 };
    for (const x of items) {
      s.by[x.status].n++;
      s.by[x.status].val += value(x);
      if (x.status !== 'NO_MOVE') s.total++;
      if (x.soq_order > 0) {
        if (x.category === 'PART') { s.nPart++; s.soqPart += x.soq_value || 0; } else { s.nBahan++; s.soqBahan += x.soq_value || 0; }
        if (x.soq_value == null) s.noPrice++;
      }
    }
    return s;
  }

  function renderCards() {
    const s = summary(state.result.items);
    const f = state.filter.status;
    const card = (status, cls, k, v, sub) =>
      `<button class="card ${cls}${f === status ? ' active' : ''}" data-status="${status}"><div class="k">${k}</div><div class="v">${v}</div><div class="s">${sub}</div></button>`;
    $('cards').innerHTML = [
      card('', '', 'Total item', num(s.total, 0), s.by.NO_MOVE.n ? `+${s.by.NO_MOVE.n} tanpa pergerakan` : ''),
      card('STOCKOUT', 'c-stockout', 'Stockout', num(s.by.STOCKOUT.n, 0), 'Stok 0, prioritas order'),
      card('ORDER', 'c-order', 'Stock Menipis', num(s.by.ORDER.n, 0), 'Stok ada tapi di bawah MIP, perlu order'),
      card('INCOMING', 'c-incoming', 'Menunggu datang', num(s.by.INCOMING.n, 0), 'Sudah tertutup PO, jangan order lagi'),
      card('OVERSTOCK', 'c-overstock', 'Overstock', num(s.by.OVERSTOCK.n, 0), 'Nilai stok ' + rpShort(s.by.OVERSTOCK.val)),
      card('NON_STOCK', 'c-nonstock', 'Non-Stock', num(s.by.NON_STOCK.n, 0), 'Pesan saat ada kebutuhan'),
      card('DEAD', 'c-dead', 'Dead stock', num(s.by.DEAD.n, 0), 'Modal tertahan ' + rpShort(s.by.DEAD.val)),
      card('SOQ', 'c-soq', 'Nilai SOQ', rpShort(s.soqPart + s.soqBahan),
        `${s.nPart + s.nBahan} item perlu order (Part ${s.nPart} · ${rpShort(s.soqPart)}, Bahan ${s.nBahan} · ${rpShort(s.soqBahan)})`
        + (s.noPrice ? ` · ${s.noPrice} item belum ada harga, belum ikut dijumlah` : '')),
    ].join('') + specialCard();
    return s;
  }

  /** Kartu pesanan khusus (ZPOL): hanya muncul bila ada yang dipantau; klik = gulir ke daftarnya. */
  function specialCard() {
    const rows = specialRows().filter((r) => r.stage !== 'DONE');
    if (!rows.length) return '';
    const waiting = rows.filter((r) => r.stage === 'ARRIVED').length, risk = rows.filter((r) => r.risk).length;
    return `<button class="card c-special" data-goto="specialPanel"><div class="k">Pesanan khusus</div><div class="v">${num(rows.length, 0)}</div>
      <div class="s">${waiting} sudah datang menunggu unit${risk ? ` · <b>${risk} berpotensi dead stock</b>` : ''}</div></button>`;
  }

  function sortVal(x, key) {
    if (key === 'priority') return RANK[x.status];
    if (key === 'status') return RANK[x.status];
    return x[key];
  }

  function visibleItems() {
    const { q, cat, status } = state.filter;
    const needle = q.trim().toUpperCase();
    const list = state.result.items.filter((x) => {
      if (cat && x.category !== cat) return false;
      if (status === 'SOQ') { if (!(x.soq_order > 0)) return false; }
      else if (status) { if (x.status !== status) return false; }
      else if (x.status === 'NO_MOVE' && !needle) return false; // disembunyikan secara default
      if (needle && !x.item_code.includes(needle) && !String(x.item_name).toUpperCase().includes(needle)
        && !String(x.supplier_name || '').toUpperCase().includes(needle)) return false;
      return true;
    });
    const { key, dir } = state.sort;
    list.sort((a, b) => {
      const va = sortVal(a, key), vb = sortVal(b, key);
      if (va == null && vb == null) return 0;
      if (va == null) return 1; // kosong selalu di bawah
      if (vb == null) return -1;
      let c = typeof va === 'string' ? va.localeCompare(vb) : va - vb;
      if (c === 0 && key === 'priority') c = (b.soq_value || 0) - (a.soq_value || 0) || b.soq_order - a.soq_order;
      return (c * dir) || a.item_code.localeCompare(b.item_code);
    });
    return list;
  }

  function rowHtml(x) {
    const st = STATUS[x.status];
    const split = x.stock > 0 && (x.stock_dms || x.stock_dps)
      ? `<small>Gudang ${num(x.stock_gudang)} · DMS ${num(x.stock_dms)} · DPS ${num(x.stock_dps)}</small>` : '';
    const merged = x.merged_from && x.merged_from.length ? `<small>termasuk ${esc(x.merged_from.join(', '))}</small>` : '';
    const sup = x.category === 'BAHAN' && x.supplier_name ? `<small>Supplier: ${esc(x.supplier_name)}</small>` : '';
    const adj = x.adjust ? `<small class="adj" title="${esc(x.adjust_reason)}">⚙ ${esc(x.adjust)}</small>` : '';
    return `<tr>
      <td class="l item"><code>${esc(x.item_code)}</code><small>${esc(x.item_name)}</small>${merged}${sup}${adj}</td>
      <td class="l"><span class="cat">${x.category}</span></td>
      <td class="l"><span class="badge ${st.cls}">${st.label}</span></td>
      <td class="num">${num(x.stock)}${split}</td>
      <td class="num">${x.on_order > 0 ? num(x.on_order) : '–'}</td>
      <td class="num">${num(x.demand)}</td>
      <td class="num">${num(x.mad, 2)}</td>
      <td class="num">${num(x.stock_level, 2)}</td>
      <td class="num">${num(x.mip)}</td>
      <td class="num${x.soq < 0 ? ' neg' : ''}">${num(x.soq)}</td>
      <td class="num strong">${x.soq_order > 0 ? num(x.soq_order) : '–'}</td>
      <td class="num">${x.soq_order > 0 ? rp(x.soq_value) : '–'}</td>
    </tr>`;
  }

  function renderTable() {
    const list = visibleItems();
    const shown = list.slice(0, state.limit);
    $('tbody').innerHTML = shown.length ? shown.map(rowHtml).join('')
      : `<tr><td colspan="${COLS.length}" class="empty">Tidak ada item yang cocok dengan pencarian/filter.</td></tr>`;
    $('count').textContent = `${num(shown.length, 0)} dari ${num(list.length, 0)} item`;
    $('btnMore').hidden = list.length <= shown.length;
    $('btnMore').textContent = `Tampilkan semua (${num(list.length, 0)} item)`;
    doc.querySelectorAll('#thead th').forEach((th) => {
      th.classList.toggle('sorted', th.dataset.key === state.sort.key);
      th.classList.toggle('desc', th.dataset.key === state.sort.key && state.sort.dir < 0);
    });
  }

  function renderDash() {
    const r = state.result;
    if (!r) return;
    $('dash').hidden = false;
    $('btnExport').disabled = false;
    $('btnSave').hidden = !((state.fresh || r.unsaved) && L.api && L.api.configured());
    const src = r.source === 'server' ? 'dari server' : 'dihitung di perangkat ini';
    $('meta').textContent = `Stok per ${tgl(r.asOf)} · demand ${tgl(r.from)} s/d ${tgl(r.asOf)} · ${src}${r.unsaved ? ' · BELUM disimpan ke server' : ''}`;
    renderCards();
    renderTable();
    renderSpecial();
  }

  function setFilter(patch) {
    Object.assign(state.filter, patch);
    state.limit = PAGE;
    $('selStatus').value = state.filter.status;
    doc.querySelectorAll('#segCat button').forEach((b) => b.classList.toggle('on', b.dataset.cat === state.filter.cat));
    renderCards();
    renderTable();
  }

  function showUpload(open) {
    $('uploadPanel').hidden = !open;
    $('btnCloseUpload').hidden = !state.result;
    if (open) { renderUpload(); root.scrollTo(0, 0); }
  }

  // ---------- Export SOQ (T6) ----------
  function soqWorkbook() {
    const r = state.result;
    const X = root.XLSX;
    const wb = X.utils.book_new();
    const sheet = (cat, title) => {
      const list = r.items.filter((x) => x.category === cat && x.soq_order > 0)
        .sort((a, b) => (cat === 'BAHAN' ? String(a.supplier_name).localeCompare(String(b.supplier_name)) : 0) || RANK[a.status] - RANK[b.status] || (b.soq_value || 0) - (a.soq_value || 0) || a.item_code.localeCompare(b.item_code));
      const head = ['No', 'Kode', 'Nama', 'Supplier', 'Status', 'Stock', 'Gudang', 'DMS', 'DPS Langkat', 'Demand 6 Bln', 'MAD', 'MIP', 'On Order', 'SOQ', 'Order', 'Harga Satuan', 'Nilai Order'];
      const rows = list.map((x, i) => [i + 1, x.item_code, x.item_name, x.supplier_name || '', STATUS[x.status].label, x.stock, x.stock_gudang, x.stock_dms, x.stock_dps,
        x.demand, Math.round(x.mad * 100) / 100, x.mip, x.on_order || 0, x.soq, x.soq_order, x.unit_price, x.soq_value]);
      const total = list.reduce((a, x) => a + (x.soq_value || 0), 0);
      const aoa = [
        [`${title} — Bengkel Daya Daihatsu Medan`],
        [`Stok per ${tgl(r.asOf)} · demand ${tgl(r.from)} s/d ${tgl(r.asOf)} · ${list.length} item`],
        [],
        head,
        ...rows,
        [],
        [null, null, 'TOTAL', null, null, null, null, null, null, null, null, null, null, null, null, null, total],
      ];
      const ws = X.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [5, 22, 38, 16, 16, 9, 9, 8, 11, 12, 9, 9, 9, 9, 9, 13, 15].map((wch) => ({ wch }));
      return ws;
    };
    X.utils.book_append_sheet(wb, sheet('PART', 'Daftar SOQ Part (ATPM)'), 'SOQ Part (ATPM)');
    X.utils.book_append_sheet(wb, sheet('BAHAN', 'Daftar SOQ Bahan'), 'SOQ Bahan');
    return wb;
  }

  function exportSOQ() {
    if (!state.result || !root.XLSX) return;
    root.XLSX.writeFile(soqWorkbook(), `SOQ_DDM_${state.result.asOf}.xlsx`);
  }

  // ---------- Server (T7, opsional) ----------
  // Formulir di halaman (bukan prompt/alert bawaan browser, yang diblokir di sebagian browser & webview).
  /** Dialog PIN serbaguna. opts: { title, hint, button, onSubmit: async (pin, user) => pesan sukses } — tanpa opts = simpan data. */
  let dlgMode = null;
  function openSaveDialog(opts) {
    dlgMode = opts && opts.onSubmit ? opts : null;
    $('saveDlgTitle').textContent = (dlgMode && dlgMode.title) || 'Simpan ke server';
    $('saveDlgHint').textContent = (dlgMode && dlgMode.hint) || 'Data hasil hitung dan daftar PO outstanding akan menggantikan data yang dilihat team.';
    $('btnDoSave').textContent = (dlgMode && dlgMode.button) || 'Simpan';
    $('pin').value = '';
    $('saveMsg').textContent = '';
    try { $('userName').value = root.localStorage.getItem('laepart.user') || ''; } catch (e) { /* abaikan */ }
    $('saveDlg').showModal();
    $('pin').focus();
  }

  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => { t.hidden = true; }, 5000);
  }

  async function saveToServer(e) {
    e.preventDefault();
    const pin = $('pin').value.trim();
    const user = $('userName').value.trim();
    if (!pin) return;
    const btn = $('btnDoSave');
    const label = btn.textContent;
    $('saveMsg').textContent = '';
    btn.disabled = true; btn.textContent = 'Menyimpan…';
    if (dlgMode) {
      try {
        const msg = await dlgMode.onSubmit(pin, user);
        try { root.localStorage.setItem('laepart.user', user); } catch (err) { /* abaikan */ }
        $('saveDlg').close();
        toast(msg || 'Tersimpan.');
      } catch (err) {
        $('saveMsg').textContent = 'Gagal menyimpan: ' + err.message;
      } finally {
        btn.disabled = false; btn.textContent = label;
      }
      return;
    }
    try {
      const p = state.parsed;
      const stock = [];
      for (const k of ['stockPart', 'stockBahan']) {
        for (const s of p[k]?.rows || []) stock.push([p[k].snapshot_date || (state.result && state.result.asOf), s.item_code, s.warehouse, s.stock]);
      }
      const log = Object.keys(SLOTS).filter((k) => p[k]).map((k) => ({ file_name: state.files[k], source: k, rows_ok: p[k].rows.length, rows_error: p[k].errors.length }));
      const payload = { pin, user };
      if (state.result) Object.assign(payload, { calc_date: state.result.asOf, items: state.result.items });
      if (state.fresh) Object.assign(payload, { stock, log });
      const sendPo = state.po.dirty || Object.keys(state.po.store).length > 0 || state.po.serverApi >= 2;
      if (sendPo) payload.po = L.po.toRows(state.po.store);
      const res = await L.api.save(payload);
      state.fresh = false;
      if (state.result && state.result.unsaved) {
        state.result.unsaved = false;
        try { root.localStorage.setItem(STORE_KEY, JSON.stringify(state.result)); } catch (err) { /* abaikan */ }
        renderDash();
      }
      try { root.localStorage.setItem('laepart.user', user); } catch (err) { /* abaikan */ }
      $('btnSave').hidden = true;
      if (sendPo && typeof res.po_saved !== 'number') {
        // backend lama: data stok tersimpan, tetapi tab PO belum dikenal
        $('saveMsg').textContent = 'Data stok tersimpan, tetapi daftar PO BELUM: backend Apps Script masih versi lama. Tempel ulang apps-script/Code.gs lalu buat versi deployment baru.';
        return;
      }
      state.po.dirty = false; state.po.serverApi = res.api || state.po.serverApi;
      persistPo();
      renderPo();
      $('saveDlg').close();
      toast('Tersimpan. Team sekarang melihat data terbaru.');
    } catch (err) {
      $('saveMsg').textContent = 'Gagal menyimpan: ' + err.message;
    } finally {
      btn.disabled = false; btn.textContent = 'Simpan';
    }
  }

  // ---------- PO outstanding (On Order) ----------
  const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  /** PO Bahan dihitung bila tanggalnya dalam 6 bulan terakhir dari tanggal data (PO lebih lama dianggap gagal/mati). */
  const poOpts = () => ({ bahanFrom: L.build.rollingFrom(state.result ? state.result.asOf : todayISO()), merge: L.build.DEFAULT_MERGE });
  const kindOf = (o) => (o.po_type === 'ZCON' || String(o.po_no).slice(4, 7) === 'SBC' ? 'BAHAN' : 'PART');

  function persistPo() {
    try { root.localStorage.setItem(PO_KEY, JSON.stringify({ store: state.po.store, dirty: state.po.dirty })); } catch (e) { /* abaikan */ }
  }

  /** Hitung ulang SOQ & status semua item dengan On Order terbaru (dari demand & stok, jadi aman dipanggil berulang). */
  function applyOnOrder() {
    if (!state.result) return;
    state.result.items = recalcWith(state.params.active);
  }

  /** Hasil hitung ulang semua item dengan parameter tertentu (untuk pratinjau di menu Parameter). */
  function recalcWith(settings) {
    return L.build.recalc(state.result.items, { onOrder: L.po.onOrderByItem(state.po.store, poOpts()).map, settings });
  }

  /** Parameter aktif baru (setelah disimpan): hitung ulang & tampilkan. */
  function setParams(settings) {
    state.params.active = L.params.clone(settings);
    try { root.localStorage.setItem(PARAM_KEY, JSON.stringify(state.params.active)); } catch (e) { /* abaikan */ }
    applyOnOrder();
    if (state.result) {
      if (state.result.source !== 'server') { try { root.localStorage.setItem(STORE_KEY, JSON.stringify(state.result)); } catch (e) { /* abaikan */ } }
      renderDash();
    }
  }

  function knownCodes() {
    const set = new Set(Object.keys(L.build.DEFAULT_MERGE));
    for (const x of (state.result && state.result.items) || []) set.add(x.item_code);
    return set;
  }

  function poChanged() {
    state.po.dirty = true;
    persistPo();
    applyOnOrder();
    if (state.result) renderDash();
    renderPo();
  }

  const ROLE = {
    COUNT: ['r-count', 'Dihitung, mengurangi SOQ'],
    NEED_DETAIL: ['r-need', 'Perlu screenshot detail PO'],
    SPECIAL: ['r-special', 'Pesanan khusus, tidak mengurangi SOQ'],
    FAILED: ['r-skip', 'Tanpa SAP Doc No., dianggap gagal'],
    OLD: ['r-skip', 'PO lama, diabaikan'],
    CLOSED: ['r-skip', 'Selesai'],
  };
  const ROLE_RANK = { NEED_DETAIL: 0, COUNT: 1, SPECIAL: 2, CLOSED: 3, FAILED: 4, OLD: 5 };

  function poCard(o, role, known) {
    const sel = (f, opts) => `<select data-f="${f}" aria-label="${f}">${opts.map((v) => `<option value="${v}"${o[f] === v ? ' selected' : ''}>${v || '–'}</option>`).join('')}</select>`;
    const [cls, label] = ROLE[role];
    const lines = (o.lines || []).map((l, i) => {
      const warn = l.corrected ? `<span class="warn">Terbaca "${esc(l.corrected)}", dikoreksi otomatis. Mohon dicek.</span>`
        : (l.item_code && state.result && !known.has(l.item_code) ? '<span class="warn">Belum ada di stok maupun penjualan 6 bulan terakhir (wajar untuk part baru / jarang). Jika kode sudah sama dengan layar ADMS, abaikan saja.</span>' : '');
      return `<tr data-line="${i}">
        <td><input class="code" data-f="item_code" value="${esc(l.item_code)}" aria-label="Kode item" spellcheck="false">${warn}</td>
        <td><input class="name" data-f="item_name" value="${esc(l.item_name || '')}" aria-label="Nama item">${l.sold ? `<span class="warn" style="color:#18794e">Terjual ${tgl(l.sold)}</span>` : ''}</td>
        <td class="q"><input class="qty" data-f="qty" inputmode="decimal" value="${esc(String(l.qty).replace('.', ','))}" aria-label="Qty"></td>
        <td><label><input type="checkbox" data-f="delivered"${l.delivered ? ' checked' : ''}> sudah datang</label></td>
        <td><button class="mini" data-act="del-line" title="Hapus baris">✕</button></td>
      </tr>`;
    }).join('');
    const body = (o.lines || []).length
      ? `<table><thead><tr><th>Kode item</th><th>Nama</th><th class="q">Qty</th><th>Kedatangan</th><th></th></tr></thead><tbody>${lines}</tbody></table>`
      : '<p class="hint" style="margin:8px 10px 0">Belum ada item. Tempel screenshot layar <b>Display Purchase Order</b> untuk PO ini, atau tambah baris manual.</p>';
    return `<div class="po" data-po="${esc(o.po_no)}">
      <div class="po-head">
        ${o.manual ? `<input class="pono" data-f="po_no" value="${esc(o.po_no)}" aria-label="No PO">` : `<code>${esc(o.po_no)}</code>`}
        <span>${tgl(o.po_date)}</span>
        ${sel('po_type', ['', 'ZPOR', 'ZPOL', 'PART', 'ZCON'])}
        ${sel('status', ['', 'Open', 'Partial', 'Completed'])}
        ${kindOf(o) === 'PART' ? `<label>SAP Doc <input class="sap" data-f="sap_doc" value="${esc(o.sap_doc || '')}" inputmode="numeric"></label>` : ''}
        ${o.pkb_no ? `<span>PKB ${esc(o.pkb_no)}</span>` : ''}
        <span class="sp"></span>
        <span class="role ${cls}">${label}</span>
        <button class="mini" data-act="del-po">Hapus PO</button>
      </div>
      ${body}
      <button class="mini" data-act="add-line" style="margin:4px 4px 6px">+ Tambah baris</button>
    </div>`;
  }

  function renderPo() {
    const P = L.po, opts = poOpts(), known = knownCodes();
    const all = Object.values(state.po.store).map((o) => ({ o, role: P.classify(o, opts) }))
      .sort((a, b) => ROLE_RANK[a.role] - ROLE_RANK[b.role] || String(b.o.po_no).localeCompare(String(a.o.po_no)));
    const active = all.filter((x) => x.role !== 'FAILED' && x.role !== 'OLD');
    const skipped = all.filter((x) => x.role === 'FAILED' || x.role === 'OLD');
    const nOut = all.filter((x) => x.role === 'COUNT' || x.role === 'NEED_DETAIL').length;
    $('poCount').hidden = !nOut;
    $('poCount').textContent = nOut;
    $('btnPoSave').hidden = !(L.api && L.api.configured() && state.po.dirty);
    $('poLog').innerHTML = state.po.log.map(([cls, msg]) => `<li class="${cls}">${esc(msg)}</li>`).join('');
    const need = all.filter((x) => x.role === 'NEED_DETAIL').length;
    const head = all.length
      ? `<p class="hint" style="margin:12px 0 0">${nOut} PO outstanding${need ? `, <b>${need} di antaranya belum punya detail item</b> sehingga belum mengurangi SOQ` : ''}.${state.po.dirty ? ' Ada perubahan yang belum disimpan ke server.' : ''}</p>`
      : '<p class="hint" style="margin:12px 0 0">Belum ada PO yang dicatat.</p>';
    const skip = skipped.length ? `<details class="skipped"><summary>${skipped.length} PO diabaikan (gagal atau lama)</summary><ul>${skipped.map(({ o, role }) =>
      `<li data-po="${esc(o.po_no)}"><code>${esc(o.po_no)}</code> · ${tgl(o.po_date)} · ${esc(o.po_type || '?')} ${esc(o.status || '')} — ${ROLE[role][1]}
        <button class="mini" data-act="del-po">Hapus</button>${role === 'FAILED' ? ' <button class="mini" data-act="unskip">Bukan gagal, hitung</button>' : ''}</li>`).join('')}</ul></details>` : '';
    $('poList').innerHTML = head + active.map(({ o, role }) => poCard(o, role, known)).join('') + skip;
  }

  const STAGE = {
    RISK: ['r-risk', 'Berpotensi dead stock'],
    ARRIVED: ['r-need', 'Sudah datang, menunggu unit'],
    ORDERED: ['r-special', 'Belum datang'],
    DONE: ['r-done', 'Selesai, sudah terpasang'],
    NEED_DETAIL: ['r-need', 'Perlu screenshot detail PO'],
  };
  function specialRows() { return L.po.specialRows(state.po.store, { today: todayISO() }); }

  function renderSpecial() {
    const rows = specialRows();
    $('specialPanel').hidden = !rows.length;
    $('specialBody').innerHTML = rows.map((r) => {
      const [cls, label] = STAGE[r.risk ? 'RISK' : r.stage];
      const age = r.age == null ? '–' : r.stage === 'ARRIVED' ? `${r.age} hari sejak tercatat datang` : r.stage === 'ORDERED' ? `${r.age} hari sejak PO` : '–';
      return `<tr><td><span class="role ${cls}">${label}</span></td>
        <td class="item"><code>${esc(r.item_code || '–')}</code><small>${esc(r.item_name || '')}</small></td>
        <td class="q">${r.qty == null ? '–' : num(r.qty)}</td><td class="q">${age}</td>
        <td><code>${esc(r.po_no)}</code><small>${tgl(r.po_date)}</small></td><td>${esc(r.pkb_no || '–')}</td><td>${esc(r.chassis_no || '–')}</td></tr>`;
    }).join('');
  }

  /** Baca screenshot PO (daftar atau detail), gabungkan ke simpanan, lalu hitung ulang. */
  async function addPoImages(fileList) {
    const files = Array.from(fileList || []).filter((f) => /^image\//.test(f.type || '') || /\.(png|jpe?g|bmp|webp)$/i.test(f.name || ''));
    if (!files.length || state.po.busy) return;
    state.po.busy = true;
    state.po.log = [];
    const match = L.po.makeMatcher([...knownCodes()]);
    const today = todayISO();
    let changed = false;
    for (let i = 0; i < files.length; i++) {
      const tag = files.length > 1 ? `Gambar ${i + 1}: ` : '';
      $('poDropText').textContent = `Membaca gambar ${i + 1} dari ${files.length}… (pemakaian pertama memuat mesin baca ±20 MB)`;
      try {
        const { lines } = await L.ocr.read(files[i]);
        const screen = L.po.detectScreen(lines);
        if (screen === 'LIST') {
          const p = L.po.parseList(lines);
          if (!p.orders.length) { state.po.log.push(['err', tag + 'Layar daftar PO dikenali, tetapi tidak ada baris PO yang terbaca.']); continue; }
          const sum = L.po.mergeParsed(state.po.store, p, today);
          state.po.log.push(['info', `${tag}Daftar PO: ${p.orders.length} PO terbaca${sum.closed ? `, ${sum.closed} sudah Completed dan ditutup` : ''}.`]);
          changed = true;
        } else if (screen === 'DETAIL') {
          const p = L.po.parseDetail(lines, { match });
          if (!p.order.po_no) { state.po.log.push(['err', tag + 'Layar detail PO dikenali, tetapi No PO tidak terbaca. Pastikan bagian atas layar ikut tertangkap.']); continue; }
          L.po.mergeParsed(state.po.store, p, today);
          state.po.log.push(['info', `${tag}Detail PO ${p.order.po_no}: ${p.order.lines.length} item terbaca.`]);
          for (const m of p.problems) state.po.log.push(['warn', tag + m]);
          changed = true;
        } else {
          state.po.log.push(['err', tag + 'Tidak dikenali sebagai layar PO ADMS (Maintain Purchase Order atau Display Purchase Order).']);
        }
      } catch (e) {
        state.po.log.push(['err', `${tag}Gagal dibaca: ${e.message || e}`]);
      }
    }
    state.po.busy = false;
    $('poDropText').textContent = 'Tempel screenshot (Ctrl+V), tarik ke sini, atau klik untuk memilih';
    if (changed) poChanged(); else renderPo();
  }

  function showPo(open) {
    $('poPanel').hidden = !open;
    if (open) { renderPo(); root.scrollTo(0, 0); }
  }

  function initPo() {
    const drop = $('poDrop'), input = $('poFile'), list = $('poList');
    $('btnPo').addEventListener('click', () => showPo($('poPanel').hidden));
    $('btnPoClose').addEventListener('click', () => showPo(false));
    $('btnPoSave').addEventListener('click', () => openSaveDialog());
    drop.addEventListener('click', () => input.click());
    drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => { if (input.files.length) addPoImages(input.files); input.value = ''; });
    for (const ev of ['dragenter', 'dragover']) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); });
    for (const ev of ['dragleave', 'drop']) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); });
    drop.addEventListener('drop', (e) => { if (e.dataTransfer.files.length) addPoImages(e.dataTransfer.files); });
    // Ctrl+V saat panel PO terbuka: ambil gambar dari clipboard (hasil Snipping Tool / Print Screen)
    doc.addEventListener('paste', (e) => {
      if ($('poPanel').hidden || !e.clipboardData) return;
      const imgs = Array.from(e.clipboardData.files || []).filter((f) => /^image\//.test(f.type));
      if (imgs.length) { e.preventDefault(); addPoImages(imgs); }
    });

    $('btnPoAdd').addEventListener('click', () => {
      let n = 1;
      while (state.po.store['BARU-' + n]) n++;
      const today = todayISO();
      state.po.store['BARU-' + n] = { po_no: 'BARU-' + n, manual: true, po_date: today, po_type: 'ZPOR', status: 'Open', sap_doc: 'MANUAL', pkb_no: '', chassis_no: '',
        kind: 'PART', has_detail: true, updated_at: today, lines: [{ item_code: '', item_name: '', qty: 0, delivered: false, delivered_at: '', sold: '' }] };
      poChanged();
    });

    list.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-act]');
      const host = e.target.closest('[data-po]');
      if (!btn || !host) return;
      const o = state.po.store[host.dataset.po];
      if (!o) return;
      const act = btn.dataset.act;
      if (act === 'del-po') delete state.po.store[o.po_no];
      else if (act === 'unskip') o.sap_doc = 'MANUAL';
      else if (act === 'add-line') { o.lines.push({ item_code: '', item_name: '', qty: 0, delivered: false, delivered_at: '', sold: '' }); o.has_detail = true; }
      else if (act === 'del-line') { o.lines.splice(Number(btn.closest('tr').dataset.line), 1); o.has_detail = o.lines.length > 0; }
      poChanged();
    });

    list.addEventListener('change', (e) => {
      const el = e.target, f = el.dataset.f, host = el.closest('[data-po]');
      if (!f || !host) return;
      const o = state.po.store[host.dataset.po];
      if (!o) return;
      const tr = el.closest('tr[data-line]');
      const today = todayISO();
      if (tr) {
        const l = o.lines[Number(tr.dataset.line)];
        if (f === 'item_code') {
          l.item_code = el.value.trim().toUpperCase(); l.corrected = '';
          // Nama kosong → ambil dari data stok/penjualan bila kodenya dikenal
          const hit = !l.item_name && state.result && state.result.items.find((x) => x.item_code === l.item_code);
          if (hit) l.item_name = hit.item_name;
        }
        else if (f === 'item_name') l.item_name = el.value.trim();
        else if (f === 'qty') { const q = Number(el.value.trim().replace(',', '.')); l.qty = Number.isFinite(q) && q >= 0 ? q : 0; }
        else if (f === 'delivered') { l.delivered = el.checked; l.delivered_at = el.checked ? (l.delivered_at || today) : ''; }
      } else if (f === 'po_no') {
        const no = el.value.trim().toUpperCase();
        if (no && no !== o.po_no && !state.po.store[no]) { delete state.po.store[o.po_no]; o.po_no = no; state.po.store[no] = o; }
        o.kind = kindOf(o);
      } else {
        o[f] = el.value.trim();
        o.kind = kindOf(o);
        if (f === 'status' && o.status === 'Completed') {
          for (const l of o.lines) { l.delivered = true; l.delivered_at = l.delivered_at || today; }
          if (o.po_type !== 'ZPOL') delete state.po.store[o.po_no]; // sudah di-GR: stok ADMS sudah naik
        }
      }
      o.updated_at = today;
      poChanged();
    });
  }

  async function loadInitial() {
    try {
      const saved = JSON.parse(root.localStorage.getItem(STORE_KEY) || 'null');
      if (saved && Array.isArray(saved.items)) {
        // hasil upload dari versi sebelum ada tanda unsaved: anggap belum disimpan (menyimpan ulang tidak merugikan)
        if (saved.source === 'upload' && saved.unsaved === undefined) saved.unsaved = !!(L.api && L.api.configured());
        state.result = saved;
      }
    } catch (e) { /* abaikan data lokal yang rusak */ }
    try {
      const ps = JSON.parse(root.localStorage.getItem(PARAM_KEY) || 'null');
      if (ps && Array.isArray(ps.suppliers) && ps.suppliers.length) state.params.active = ps;
    } catch (e) { /* abaikan */ }
    try {
      const po = JSON.parse(root.localStorage.getItem(PO_KEY) || 'null');
      if (po && po.store) { state.po.store = po.store; state.po.dirty = !!po.dirty; }
    } catch (e) { /* abaikan */ }
    applyOnOrder();
    if (state.result) renderDash(); else showUpload(true);
    renderPo();

    if (L.api && L.api.configured()) {
      try {
        const r = await L.api.fetchResults();
        state.po.serverApi = r.api || 0;
        state.params.serverApi = r.api || 0;
        if (r.settings) {
          // backend v3: parameter di server berlaku untuk semua (tab kosong = bawaan)
          const ps = L.params.fromRows(r.settings);
          state.params.active = ps || L.params.clone(L.params.DEFAULT_SETTINGS);
          state.params.fromServer = !!ps;
          state.params.log = Array.isArray(r.param_log) ? r.param_log : [];
          try { root.localStorage.setItem(PARAM_KEY, JSON.stringify(state.params.active)); } catch (e) { /* abaikan */ }
        }
        // PO dari server dipakai kecuali ada perubahan lokal yang belum disimpan
        if (Array.isArray(r.po) && !state.po.dirty) { state.po.store = L.po.fromRows(r.po); persistPo(); }
        // data server dipakai, kecuali hasil upload di perangkat ini belum disimpan dan tidak lebih lama dari data server
        const keepLocal = state.fresh || (state.result && state.result.unsaved && !(r.calc_date > state.result.asOf));
        if (r && r.items && r.items.length && !keepLocal) {
          state.result = { asOf: r.calc_date, from: L.build.rollingFrom(r.calc_date), items: r.items, source: 'server' };
          showUpload(false);
        }
        applyOnOrder();
        if (state.result) renderDash();
        renderPo();
        if (L.paramUi) L.paramUi.serverLoaded();
      } catch (e) {
        if (!state.result) { state.fileErrors = ['Data dari server tidak bisa dimuat: ' + e.message]; renderUpload(); }
      }
    }
  }

  function init() {
    $('thead').innerHTML = COLS.map((c) => `<th class="${c.left ? 'l' : ''}" data-key="${c.key}">${c.label}</th>`).join('');
    $('selStatus').innerHTML = '<option value="">Semua status</option><option value="SOQ">Masuk daftar order (SOQ &gt; 0)</option>'
      + Object.entries(STATUS).map(([k, v]) => `<option value="${k}">${v.label}</option>`).join('');

    const drop = $('drop'), input = $('fileInput');
    drop.addEventListener('click', () => input.click());
    drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    input.addEventListener('change', () => { if (input.files.length) addFiles(input.files); input.value = ''; });
    for (const ev of ['dragenter', 'dragover']) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); });
    for (const ev of ['dragleave', 'drop']) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); });
    drop.addEventListener('drop', (e) => { if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files); });
    // file yang meleset dari kotak jangan sampai dibuka browser
    for (const ev of ['dragover', 'drop']) root.addEventListener(ev, (e) => e.preventDefault());

    $('btnUpload').addEventListener('click', () => showUpload($('uploadPanel').hidden));
    $('btnCloseUpload').addEventListener('click', () => showUpload(false));
    $('btnReset').addEventListener('click', resetUpload);
    $('btnCalc').addEventListener('click', compute);
    $('btnExport').addEventListener('click', exportSOQ);
    $('btnSave').addEventListener('click', () => openSaveDialog());
    $('saveForm').addEventListener('submit', saveToServer);
    $('btnCancelSave').addEventListener('click', () => $('saveDlg').close());
    $('btnMore').addEventListener('click', () => { state.limit = Infinity; renderTable(); });
    initPo();

    let t;
    $('q').addEventListener('input', (e) => { clearTimeout(t); t = setTimeout(() => setFilter({ q: e.target.value }), 150); });
    $('selStatus').addEventListener('change', (e) => setFilter({ status: e.target.value }));
    $('segCat').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) setFilter({ cat: b.dataset.cat }); });
    $('cards').addEventListener('click', (e) => {
      const b = e.target.closest('.card');
      if (!b) return;
      if (b.dataset.goto) $(b.dataset.goto).scrollIntoView({ behavior: 'smooth', block: 'start' });
      else setFilter({ status: b.dataset.status });
    });
    $('thead').addEventListener('click', (e) => {
      const th = e.target.closest('th');
      if (!th) return;
      const key = th.dataset.key;
      // kolom teks mulai A→Z, kolom angka mulai dari terbesar
      const first = ['item_code', 'category', 'status'].includes(key) ? 1 : -1;
      state.sort = state.sort.key === key ? { key, dir: -state.sort.dir } : { key, dir: first };
      renderTable();
    });

    loadInitial();
  }

  L.ui = { addFiles, compute, exportSOQ, soqWorkbook, addPoImages, state, recalcWith, setParams, openSaveDialog, toast, summary, todayISO,
    fmt: { esc, num, rp, rpShort, tgl }, STATUS };
  if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', init); else init();
})(window);
