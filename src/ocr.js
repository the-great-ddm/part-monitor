/**
 * Daya-LAE Part Monitor — OCR screenshot di browser (PaddleOCR lewat ppu-paddle-ocr + onnxruntime-web).
 * Gambar diproses di perangkat ini saja; tidak ada yang dikirim ke server.
 *
 * Mesin (.wasm) dan model disimpan di src/vendor/ agar tidak bergantung pada server lain;
 * pustaka JS kecilnya dimuat dari CDN dengan versi terkunci. Lihat <script type="importmap"> di index.html.
 * Dimuat hanya saat pertama kali dipakai (±20 MB, sekali per browser).
 */
(function (root) {
  'use strict';
  const LIB = 'https://cdn.jsdelivr.net/npm/ppu-paddle-ocr@6.7.1/web/index.js';
  const SCALE = 2; // layar ADMS berhuruf ±11 px; diperbesar 2× sebelum dibaca
  let ready = null;

  function service() {
    if (!ready) {
      ready = (async () => {
        const m = await import(LIB);
        const base = new URL('src/vendor/ocr-models/', root.document.baseURI).href;
        const svc = new m.PaddleOcrService({
          model: { detection: base + 'PP-OCRv6_tiny_det.ort', recognition: base + 'PP-OCRv6_tiny_rec.ort', charactersDictionary: base + 'ppocrv6_tiny_dict.txt' },
          session: { executionProviders: ['wasm'], graphOptimizationLevel: 'all' },
        });
        await svc.initialize();
        return svc;
      })();
      ready.catch(() => { ready = null; }); // boleh dicoba lagi setelah gagal (mis. internet putus)
    }
    return ready;
  }

  /** "Tinta" sel = jumlah piksel abu-abu; dipakai po.js untuk membedakan kotak centang kosong vs tercentang. */
  function inkOf(ctx, b, w, h) {
    const x = Math.max(0, Math.floor(b.x)), y = Math.max(0, Math.floor(b.y));
    const cw = Math.min(w - x, Math.ceil(b.width)), ch = Math.min(h - y, Math.ceil(b.height));
    if (cw <= 0 || ch <= 0) return 0;
    const d = ctx.getImageData(x, y, cw, ch).data;
    let mid = 0;
    for (let i = 0; i < d.length; i += 4) {
      const lum = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
      if (lum >= 90 && lum < 170) mid++;
    }
    return mid;
  }

  /** blob gambar → { lines: [[{ text, box, ink? }]], scale } */
  async function read(blob) {
    const svc = await service();
    const bmp = await root.createImageBitmap(blob);
    const canvas = root.document.createElement('canvas');
    canvas.width = bmp.width * SCALE;
    canvas.height = bmp.height * SCALE;
    // PENTING: kanvas biasa (bukan willReadFrequently). Kanvas perangkat lunak memperbesar gambar dengan cara lain
    // dan terbukti membuat nomor PO salah terbaca pada contoh layar ADMS (uji 17 Nov 2026).
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
    const res = await svc.recognize(canvas);
    // salinan 1:1 khusus untuk membaca piksel kotak centang
    const copy = root.document.createElement('canvas');
    copy.width = canvas.width;
    copy.height = canvas.height;
    const rctx = copy.getContext('2d', { willReadFrequently: true });
    rctx.drawImage(canvas, 0, 0);
    const lines = (res.lines || []).map((row) => row.map((c) => {
      const cell = { text: c.text || '', box: c.box };
      if (cell.text.trim().length <= 2) cell.ink = inkOf(rctx, c.box, copy.width, copy.height);
      return cell;
    }));
    return { lines, scale: SCALE };
  }

  root.LAEPart = Object.assign(root.LAEPart || {}, { ocr: { read } });
})(window);
