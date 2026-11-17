/**
 * Daya-LAE Part Monitor — panggilan ke Apps Script Web App (apps-script/Code.gs).
 * Opsional: tanpa src/config.js aplikasi tetap jalan lokal di browser.
 */
(function (root) {
  'use strict';
  const url = () => ((root.LAEPart && root.LAEPart.config && root.LAEPart.config.apiUrl) || '').trim();
  const configured = () => /^https:\/\//.test(url());

  async function read(res) {
    if (!res.ok) throw new Error('Server menjawab ' + res.status);
    let data;
    try { data = await res.json(); } catch (e) { throw new Error('Jawaban server tidak terbaca. Periksa URL & izin akses Web App.'); }
    if (!data.ok) throw new Error(data.error || 'Permintaan ditolak server');
    return data;
  }

  /** → { ok, calc_date, items: [...] } */
  async function fetchResults() {
    return read(await root.fetch(url() + '?action=results'));
  }

  /** payload: { pin, user, calc_date, items, stock: [[snapshot_date, item_code, warehouse, stock]], log } */
  async function save(payload) {
    // text/plain agar browser tidak mengirim preflight CORS (Apps Script tidak melayani OPTIONS)
    return read(await root.fetch(url(), {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'save', ...payload }),
    }));
  }

  root.LAEPart = Object.assign(root.LAEPart || {}, { api: { configured, fetchResults, save } });
})(window);
