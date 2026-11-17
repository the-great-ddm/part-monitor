// Salin file ini menjadi src/config.js (di-.gitignore), lalu isi URL Web App Apps Script (berakhiran /exec).
// Kosongkan apiUrl untuk memakai aplikasi secara lokal saja.
// PIN TIDAK disimpan di sini — PIN diatur di Apps Script (Script Properties → PIN) dan diketik saat menyimpan.
window.LAEPart = Object.assign(window.LAEPart || {}, {
  config: {
    apiUrl: '',
  },
});
