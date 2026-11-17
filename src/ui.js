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
    ORDER: { label: 'Perlu Order', cls: 's-order' },
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
      r = L.build.buildResults(state.parsed);
    } catch (e) {
      state.fileErrors = ['Perhitungan gagal: ' + e.message];
      renderUpload();
      return;
    }
    state.result = { asOf: r.asOf, from: r.from, items: r.items, source: 'upload', files: { ...state.files } };
    state.fresh = true;
    try { root.localStorage.setItem(STORE_KEY, JSON.stringify(state.result)); } catch (e) { /* penyimpanan lokal tidak tersedia */ }
    showUpload(false);
    renderDash();
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
      card('ORDER', 'c-order', 'Perlu Order', num(s.by.ORDER.n, 0), 'Stok di bawah MIP'),
      card('OVERSTOCK', 'c-overstock', 'Overstock', num(s.by.OVERSTOCK.n, 0), 'Nilai stok ' + rpShort(s.by.OVERSTOCK.val)),
      card('NON_STOCK', 'c-nonstock', 'Non-Stock', num(s.by.NON_STOCK.n, 0), 'Pesan saat ada kebutuhan'),
      card('DEAD', 'c-dead', 'Dead stock', num(s.by.DEAD.n, 0), 'Modal tertahan ' + rpShort(s.by.DEAD.val)),
      card('SOQ', 'c-soq', 'Nilai SOQ', rpShort(s.soqPart + s.soqBahan),
        `Part ${s.nPart} item ${rpShort(s.soqPart)} · Bahan ${s.nBahan} item ${rpShort(s.soqBahan)}`
        + (s.noPrice ? ` · ${s.noPrice} item belum ada harga` : '')),
    ].join('');
    return s;
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
      if (needle && !x.item_code.includes(needle) && !String(x.item_name).toUpperCase().includes(needle)) return false;
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
    return `<tr>
      <td class="l item"><code>${esc(x.item_code)}</code><small>${esc(x.item_name)}</small>${merged}</td>
      <td class="l"><span class="cat">${x.category}</span></td>
      <td class="l"><span class="badge ${st.cls}">${st.label}</span></td>
      <td class="num">${num(x.stock)}${split}</td>
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
    $('btnSave').hidden = !(state.fresh && L.api && L.api.configured());
    const src = r.source === 'server' ? 'dari server' : 'dihitung di perangkat ini';
    $('meta').textContent = `Stok per ${tgl(r.asOf)} · demand ${tgl(r.from)} s/d ${tgl(r.asOf)} · ${src}`;
    renderCards();
    renderTable();
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
        .sort((a, b) => RANK[a.status] - RANK[b.status] || (b.soq_value || 0) - (a.soq_value || 0) || a.item_code.localeCompare(b.item_code));
      const head = ['No', 'Kode', 'Nama', 'Status', 'Stock', 'Gudang', 'DMS', 'DPS Langkat', 'Demand 6 Bln', 'MAD', 'MIP', 'SOQ', 'Order', 'Harga Satuan', 'Nilai Order'];
      const rows = list.map((x, i) => [i + 1, x.item_code, x.item_name, STATUS[x.status].label, x.stock, x.stock_gudang, x.stock_dms, x.stock_dps,
        x.demand, Math.round(x.mad * 100) / 100, x.mip, x.soq, x.soq_order, x.unit_price, x.soq_value]);
      const total = list.reduce((a, x) => a + (x.soq_value || 0), 0);
      const aoa = [
        [`${title} — Bengkel Daya Daihatsu Medan`],
        [`Stok per ${tgl(r.asOf)} · demand ${tgl(r.from)} s/d ${tgl(r.asOf)} · ${list.length} item`],
        [],
        head,
        ...rows,
        [],
        [null, null, 'TOTAL', null, null, null, null, null, null, null, null, null, null, null, total],
      ];
      const ws = X.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [5, 22, 38, 12, 9, 9, 8, 11, 12, 9, 9, 9, 9, 13, 15].map((wch) => ({ wch }));
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
  function openSaveDialog() {
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
    $('saveMsg').textContent = '';
    btn.disabled = true; btn.textContent = 'Menyimpan…';
    try {
      const p = state.parsed;
      const stock = [];
      for (const k of ['stockPart', 'stockBahan']) {
        for (const s of p[k]?.rows || []) stock.push([p[k].snapshot_date || state.result.asOf, s.item_code, s.warehouse, s.stock]);
      }
      const log = Object.keys(SLOTS).filter((k) => p[k]).map((k) => ({ file_name: state.files[k], source: k, rows_ok: p[k].rows.length, rows_error: p[k].errors.length }));
      await L.api.save({ pin, user, calc_date: state.result.asOf, items: state.result.items, stock, log });
      state.fresh = false;
      try { root.localStorage.setItem('laepart.user', user); } catch (err) { /* abaikan */ }
      $('saveDlg').close();
      $('btnSave').hidden = true;
      toast('Tersimpan. Team sekarang melihat data terbaru.');
    } catch (err) {
      $('saveMsg').textContent = 'Gagal menyimpan: ' + err.message;
    } finally {
      btn.disabled = false; btn.textContent = 'Simpan';
    }
  }

  async function loadInitial() {
    try {
      const saved = JSON.parse(root.localStorage.getItem(STORE_KEY) || 'null');
      if (saved && Array.isArray(saved.items)) state.result = saved;
    } catch (e) { /* abaikan data lokal yang rusak */ }
    if (state.result) renderDash(); else showUpload(true);

    if (L.api && L.api.configured()) {
      try {
        const r = await L.api.fetchResults();
        if (r && r.items && r.items.length && !state.fresh) {
          state.result = { asOf: r.calc_date, from: L.build.rollingFrom(r.calc_date), items: r.items, source: 'server' };
          showUpload(false);
          renderDash();
        }
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
    $('btnSave').addEventListener('click', openSaveDialog);
    $('saveForm').addEventListener('submit', saveToServer);
    $('btnCancelSave').addEventListener('click', () => $('saveDlg').close());
    $('btnMore').addEventListener('click', () => { state.limit = Infinity; renderTable(); });

    let t;
    $('q').addEventListener('input', (e) => { clearTimeout(t); t = setTimeout(() => setFilter({ q: e.target.value }), 150); });
    $('selStatus').addEventListener('change', (e) => setFilter({ status: e.target.value }));
    $('segCat').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) setFilter({ cat: b.dataset.cat }); });
    $('cards').addEventListener('click', (e) => { const b = e.target.closest('.card'); if (b) setFilter({ status: b.dataset.status }); });
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

  L.ui = { addFiles, compute, exportSOQ, soqWorkbook, state };
  if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', init); else init();
})(window);
