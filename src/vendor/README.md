# Berkas pihak ketiga untuk OCR (dipakai `src/ocr.js`)

Disimpan di sini agar situs tidak bergantung pada server lain saat membaca screenshot PO.
Jangan diubah; untuk memperbarui, ganti dengan berkas dari sumber resminya dan uji ulang dengan `tests/po.test.js` + screenshot contoh.

| Folder | Isi | Sumber |
|--------|-----|--------|
| `ort/` | onnxruntime-web 1.26.0 (`ort.wasm.bundle.min.mjs`, `ort-wasm-simd-threaded.wasm`, `.mjs`) | npm `onnxruntime-web` — Microsoft, lisensi MIT |
| `ocr-models/` | Model PP-OCRv6 tiny (deteksi + pengenalan) dan kamusnya | PaddleOCR (PaddlePaddle), lisensi Apache-2.0, dikemas oleh proyek `ppu-paddle-ocr` (Hugging Face `snowfluke/ppu-paddle-ocr-models`) |

Pustaka JS `ppu-paddle-ocr@6.7.1` dan `ppu-ocv@4.0.1` dimuat dari CDN jsDelivr dengan versi terkunci (lihat `<script type="importmap">` di `index.html`).
