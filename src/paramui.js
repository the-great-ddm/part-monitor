/**
 * Daya-LAE Part Monitor — menu "Parameter": atur rumus per supplier & penyesuaian per item (aturan: docs/FORMULAS.md §6).
 * Perubahan dikumpulkan sebagai draf → pratinjau dampak → simpan dengan PIN (berlaku untuk semua) → riwayat.
 * Perhitungan ada di params.js / calc.js / build.js; file ini hanya tampilan.
 */
(function (root) {
  'use strict';
  const L = root.LAEPart;
  const P = L.params;
  const doc = root.document;
  const $ = (id) => doc.getElementById(id);
  const U = L.ui;
  const { esc, num, rpShort, tgl } = U.fmt;
  const DRAFT_KEY = 'laepart.params.draft';

  const st = { draft: null };
  const active = () => U.state.params.active;
  const changes = () => (st.draft ? P.diff(active(), st.draft) : []);

  function persistDraft() {
    try {
      if (st.draft && changes().length) root.localStorage.setItem(DRAFT_KEY, JSON.stringify(st.draft));
      else root.localStorage.removeItem(DRAFT_KEY);
    } catch (e) { /* abaikan */ }
  }

  // ---------- Supplier ----------
  const numIn = (f, v, title, step = '1') => `<input class="n" type="number" min="0" step="${step}" data-f="${f}" value="${esc(v)}" aria-label="${title}" title="${title}">`;

  function supRow(s, i) {
    const opt = (v, l, cur) => `<option value="${v}"${cur === v ? ' selected' : ''}>${l}</option>`;
    return `<tr data-i="${i}">
      <td><input class="t" data-f="name" value="${esc(s.name)}" aria-label="Nama supplier"></td>
      <td><input class="t s" data-f="city" value="${esc(s.city)}" aria-label="Kota" placeholder="Medan / Jakarta"></td>
      <td><select data-f="category" aria-label="Kategori">${opt('PART', 'Part', s.category)}${opt('BAHAN', 'Bahan', s.category)}</select></td>
      <td><input class="t code" data-f="prefixes" value="${esc((s.prefixes || []).join(', '))}" placeholder="(bawaan)" aria-label="Awalan kode" spellcheck="false"></td>
      <td>${numIn('lead', s.lead, 'Lead time (hari)')}</td>
      <td>${numIn('cycle', s.cycle, 'Siklus order (hari)')}</td>
      <td>${numIn('safety', s.safety, 'Safety stock (hari)')}</td>
      <td class="calc" data-calc="level"></td>
      <td><select data-f="level_round" aria-label="Pembulatan Stock Level" title="Pembulatan Stock Level">${opt('UP1', 'SL naik 0,1', s.level_round)}${opt('NONE', 'SL tanpa', s.level_round)}</select>
        <select data-f="mip_round" aria-label="Pembulatan MIP" title="Pembulatan MIP">${opt('ROUND0', 'MIP bulat', s.mip_round)}${opt('ROUNDUP1', 'MIP naik 0,1', s.mip_round)}</select></td>
      <td>${numIn('factor', s.factor, 'Pengali MIP (%)')}</td>
      <td>${numIn('multiple', s.multiple, 'Kelipatan order (0 = tidak dibulatkan)', 'any')}</td>
      <td>${numIn('tol_days', s.tol_days, 'Toleransi overstock (hari pemakaian)')}</td>
      <td class="calc" data-calc="count"></td>
      <td><button class="mini" data-act="del-sup" title="Hapus supplier">✕</button></td>
    </tr>`;
  }

  // ---------- Penyesuaian item ----------
  function ruleRow(r, i) {
    const sups = st.draft.suppliers.map((s) => `<option value="${esc(s.id)}"${r.supplier === s.id ? ' selected' : ''}>${esc(s.name)}</option>`).join('');
    const opt = (f, title, step = 'any') => `<input class="n" type="number" min="0" step="${step}" data-f="${f}" value="${r[f] == null ? '' : esc(r[f])}" aria-label="${title}" title="${title}" placeholder="–">`;
    return `<tr data-i="${i}">
      <td><input class="t code" data-f="key" value="${esc(r.key)}" list="paramCodes" placeholder="Kode atau AWALAN*" aria-label="Kode item" spellcheck="false">
        <small data-calc="match"></small></td>
      <td><select data-f="supplier" aria-label="Supplier"><option value="">(otomatis)</option>${sups}</select></td>
      <td>${opt('mip', 'MIP manual')}</td>
      <td>${opt('level', 'Stock Level manual (bulan)')}</td>
      <td>${opt('factor', 'Pengali MIP (%)', '1')}</td>
      <td>${opt('multiple', 'Kelipatan order')}</td>
      <td><input class="t" data-f="reason" value="${esc(r.reason)}" placeholder="Wajib, mis. instruksi management" aria-label="Alasan"></td>
      <td><small>${r.by ? esc(r.by) + '<br>' : ''}${r.at ? esc(tgl(r.at)) : ''}</small></td>
      <td><button class="mini" data-act="del-rule" title="Hapus penyesuaian">✕</button></td>
    </tr>`;
  }

  function renderTables() {
    $('supBody').innerHTML = st.draft.suppliers.map(supRow).join('');
    $('ruleBody').innerHTML = st.draft.rules.length ? st.draft.rules.map(ruleRow).join('')
      : '<tr><td colspan="9" class="empty">Belum ada. Contoh: MIP manual untuk part tertentu, atau <code>D90000-*</code> pengali 120% untuk semua oli.</td></tr>';
    const items = (U.state.result && U.state.result.items) || [];
    $('paramCodes').innerHTML = items.slice(0, 2000).map((x) => `<option value="${esc(x.item_code)}">${esc(x.item_name)}</option>`).join('');
    renderCalc();
  }

  // ---------- Angka turunan, pratinjau, kesalahan (tanpa menggambar ulang isian agar fokus tidak hilang) ----------
  function renderCalc() {
    const d = st.draft;
    const items = (U.state.result && U.state.result.items) || [];
    const after = items.length ? U.recalcWith(d) : [];
    const per = {};
    for (const x of after) per[x.supplier] = (per[x.supplier] || 0) + 1;
    doc.querySelectorAll('#supBody tr').forEach((tr) => {
      const s = d.suppliers[Number(tr.dataset.i)];
      if (!s) return;
      tr.querySelector('[data-calc="level"]').innerHTML = `<b>${num(P.supplierLevel(s), 4)}</b> bln`;
      tr.querySelector('[data-calc="count"]').textContent = items.length ? num(per[s.id] || 0, 0) : '–';
    });
    const byCode = Object.fromEntries(items.map((x) => [x.item_code, x]));
    doc.querySelectorAll('#ruleBody tr[data-i]').forEach((tr) => {
      const r = d.rules[Number(tr.dataset.i)];
      const k = String(r.key || '').trim().toUpperCase();
      let t = '';
      if (k.endsWith('*')) { const pre = k.slice(0, -1); t = `${num(items.filter((x) => x.item_code.startsWith(pre)).length, 0)} item cocok`; }
      else if (k) t = byCode[k] ? byCode[k].item_name : 'Kode tidak ada di data saat ini';
      tr.querySelector('[data-calc="match"]').textContent = t;
    });

    const errs = P.validate(d);
    const ch = changes();
    $('paramErrs').innerHTML = errs.map((m) => `<li class="err">${esc(m)}</li>`).join('');
    $('btnParamSave').disabled = !ch.length || errs.length > 0;
    $('btnParamUndo').disabled = !ch.length;
    $('paramState').textContent = ch.length ? `${ch.length} perubahan belum disimpan.` : 'Tidak ada perubahan.';
    renderPreview(items, after, ch);
    persistDraft();
  }

  function renderPreview(before, after, ch) {
    const box = $('paramPreview');
    if (!ch.length) { box.innerHTML = ''; return; }
    const list = `<ul class="chg">${ch.map((c) => `<li>${esc(c)}</li>`).join('')}</ul>`;
    if (!before.length) { box.innerHTML = `<h3>Perubahan</h3>${list}<p class="hint">Belum ada data stok, dampak akan terlihat setelah upload.</p>`; return; }
    const a = U.summary(before), b = U.summary(after);
    const bm = Object.fromEntries(before.map((x) => [x.item_code, x]));
    const changed = after.filter((x) => {
      const o = bm[x.item_code];
      return o && (o.mip !== x.mip || o.soq_order !== x.soq_order || o.status !== x.status || o.stock_level !== x.stock_level);
    }).sort((x, y) => Math.abs((y.soq_value || 0) - (bm[y.item_code].soq_value || 0)) - Math.abs((x.soq_value || 0) - (bm[x.item_code].soq_value || 0)));
    const arrow = (o, n, f = (v) => num(v)) => (o === n ? f(n) : `${f(o)} → <b>${f(n)}</b>`);
    const stat = Object.keys(U.STATUS).filter((k) => a.by[k].n !== b.by[k].n)
      .map((k) => `${U.STATUS[k].label} ${a.by[k].n} → <b>${b.by[k].n}</b>`).join(' · ');
    const rows = changed.slice(0, 100).map((x) => {
      const o = bm[x.item_code];
      return `<tr><td class="l item"><code>${esc(x.item_code)}</code><small>${esc(x.item_name)}</small></td>
        <td class="l">${esc(x.supplier_name)}</td>
        <td class="num">${arrow(o.stock_level, x.stock_level, (v) => num(v, 2))}</td>
        <td class="num">${arrow(o.mip, x.mip)}</td>
        <td class="num">${arrow(o.soq_order, x.soq_order)}</td>
        <td class="l">${o.status === x.status ? U.STATUS[x.status].label : `${U.STATUS[o.status].label} → <b>${U.STATUS[x.status].label}</b>`}</td></tr>`;
    }).join('');
    const nA = a.nPart + a.nBahan, nB = b.nPart + b.nBahan;
    box.innerHTML = `<h3>Pratinjau dampak</h3>${list}
      <p class="impact"><b>${num(changed.length, 0)}</b> item berubah · Item perlu order ${arrow(nA, nB, (v) => num(v, 0))}
        · Nilai SOQ ${arrow(a.soqPart + a.soqBahan, b.soqPart + b.soqBahan, rpShort)}${stat ? ` · ${stat}` : ''}</p>
      ${changed.length ? `<div class="tablewrap"><table class="prev"><thead><tr><th class="l">Item</th><th class="l">Supplier</th><th>Stock Level</th><th>MIP</th><th>Order</th><th class="l">Status</th></tr></thead>
        <tbody>${rows}</tbody></table></div>${changed.length > 100 ? `<p class="hint">Menampilkan 100 dari ${num(changed.length, 0)} item (urut dampak nilai terbesar).</p>` : ''}` : ''}`;
  }

  function renderLog() {
    const log = U.state.params.log || [];
    $('paramLog').innerHTML = log.length
      ? log.map((l) => `<li><span class="when">${esc(l.timestamp)}</span> <b>${esc(l.user || '–')}</b> — ${esc(l.change)}</li>`).join('')
      : '<li class="hint">Belum ada perubahan tercatat.</li>';
    const p = U.state.params;
    const configured = L.api && L.api.configured();
    $('paramSrc').textContent = !configured ? 'Server belum diatur: parameter hanya berlaku di perangkat ini.'
      : p.serverApi && p.serverApi < 3 ? 'Backend Apps Script masih versi lama: tempel ulang apps-script/Code.gs lalu buat versi deployment baru agar parameter bisa disimpan.'
        : p.fromServer ? 'Parameter dibaca dari server dan berlaku untuk semua yang membuka situs.' : 'Memakai parameter bawaan (rumus Excel). Belum ada perubahan yang disimpan di server.';
  }

  function render() {
    if (!st.draft) st.draft = P.clone(active());
    renderTables();
    renderLog();
  }

  function open(show) {
    $('paramPanel').hidden = !show;
    if (show) { render(); root.scrollTo(0, 0); }
  }

  // ---------- Simpan ----------
  function stamp(user) {
    const today = U.todayISO();
    const old = Object.fromEntries(active().rules.map((r) => [String(r.key).trim().toUpperCase(), JSON.stringify({ ...r, by: '', at: '' })]));
    for (const r of st.draft.rules) {
      r.key = String(r.key || '').trim().toUpperCase();
      if (old[r.key] !== JSON.stringify({ ...r, by: '', at: '' })) { r.by = user || '–'; r.at = today; }
    }
  }

  function save() {
    const ch = changes();
    if (!ch.length || P.validate(st.draft).length) return;
    const configured = L.api && L.api.configured();
    if (!configured) {
      stamp('');
      U.setParams(st.draft);
      st.draft = P.clone(active());
      render();
      U.toast('Parameter diterapkan di perangkat ini.');
      return;
    }
    U.openSaveDialog({
      title: 'Simpan parameter',
      hint: `${ch.length} perubahan akan berlaku untuk semua yang membuka situs dan dicatat di riwayat.`,
      button: 'Simpan & terapkan',
      onSubmit: async (pin, user) => {
        if (!user) throw new Error('isi nama Anda untuk riwayat perubahan.');
        stamp(user);
        let res;
        try {
          res = await L.api.saveParams({ pin, user, settings: P.toRows(st.draft), changes: ch });
        } catch (e) {
          if (/Aksi tidak dikenal/.test(e.message)) throw new Error('backend Apps Script masih versi lama. Tempel ulang apps-script/Code.gs lalu Deploy → Manage deployments → Edit → New version.');
          throw e;
        }
        if (!res.params_saved) throw new Error('server tidak mengonfirmasi penyimpanan.');
        const when = new Date().toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'short' });
        U.state.params.log = [...ch.map((c) => ({ timestamp: when, user, change: c })), ...(U.state.params.log || [])].slice(0, 50);
        U.state.params.fromServer = true;
        U.state.params.serverApi = res.api || U.state.params.serverApi;
        U.setParams(st.draft);
        st.draft = P.clone(active());
        render();
        return 'Parameter tersimpan dan diterapkan. Dashboard sudah dihitung ulang.';
      },
    });
  }

  // ---------- Isian ----------
  function onInput(e) {
    const el = e.target, f = el.dataset.f, tr = el.closest('tr[data-i]');
    if (!f || !tr) return;
    const i = Number(tr.dataset.i);
    const NUM = ['lead', 'cycle', 'safety', 'factor', 'multiple', 'tol_days'];
    if (tr.closest('#supBody')) {
      const s = st.draft.suppliers[i];
      if (f === 'prefixes') s.prefixes = el.value.split(',').map((p) => p.trim().toUpperCase()).filter(Boolean);
      else if (NUM.includes(f)) s[f] = el.value === '' ? 0 : Number(el.value);
      else s[f] = el.value;
      if (f === 'name' || f === 'category') { renderTables(); return; } // nama supplier dipakai di pilihan item
    } else {
      const r = st.draft.rules[i];
      if (['mip', 'level', 'factor', 'multiple'].includes(f)) r[f] = el.value === '' ? null : Number(el.value);
      else r[f] = f === 'key' ? el.value.trim().toUpperCase() : el.value;
    }
    renderCalc();
  }

  function init() {
    $('btnParam').addEventListener('click', () => open($('paramPanel').hidden));
    $('btnParamClose').addEventListener('click', () => open(false));
    $('btnParamSave').addEventListener('click', save);
    $('btnParamUndo').addEventListener('click', () => { st.draft = P.clone(active()); render(); });
    $('btnParamDefault').addEventListener('click', () => { st.draft = P.clone(P.DEFAULT_SETTINGS); render(); });
    $('btnAddSup').addEventListener('click', () => {
      let n = 1;
      while (st.draft.suppliers.some((s) => s.id === 'SUP' + n)) n++;
      st.draft.suppliers.push({ id: 'SUP' + n, name: 'Supplier baru', city: '', category: 'BAHAN', prefixes: [], lead: 7, cycle: 7, safety: 15,
        level_round: 'NONE', mip_round: 'ROUNDUP1', factor: 100, multiple: 0, tol_days: 0 });
      renderTables();
    });
    $('btnAddRule').addEventListener('click', () => {
      st.draft.rules.push({ key: '', supplier: '', mip: null, level: null, factor: null, multiple: null, reason: '', by: '', at: '' });
      renderTables();
      const inputs = doc.querySelectorAll('#ruleBody input[data-f="key"]');
      if (inputs.length) inputs[inputs.length - 1].focus();
    });
    for (const id of ['supBody', 'ruleBody']) {
      $(id).addEventListener('change', onInput);
      $(id).addEventListener('click', (e) => {
        const b = e.target.closest('[data-act]');
        if (!b) return;
        const i = Number(b.closest('tr').dataset.i);
        if (b.dataset.act === 'del-sup') {
          const id2 = st.draft.suppliers[i].id;
          st.draft.suppliers.splice(i, 1);
          for (const r of st.draft.rules) if (r.supplier === id2) r.supplier = '';
        } else st.draft.rules.splice(i, 1);
        renderTables();
      });
    }
    try {
      const d = JSON.parse(root.localStorage.getItem(DRAFT_KEY) || 'null');
      if (d && Array.isArray(d.suppliers)) st.draft = d;
    } catch (e) { /* abaikan */ }
  }

  L.paramUi = {
    /** Dipanggil ui.js setelah parameter dari server dimuat. */
    serverLoaded() {
      if (st.draft && !changes().length) st.draft = null; // tidak ada draf: ikuti server
      if (!$('paramPanel').hidden) render(); else renderLogOnly();
    },
    state: st,
  };
  function renderLogOnly() { if (!$('paramPanel').hidden) renderLog(); }

  if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', init); else init();
})(window);
