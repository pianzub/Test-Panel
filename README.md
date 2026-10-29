# Bakerzin Food Test Panel — GitHub Pages

Paket ini menggunakan frontend `index.html` dari versi sumber 29 Oktober 2026 dan backend Google Apps Script pada `Code.gs`.

## File dalam paket

- `index.html` — frontend untuk GitHub Pages.
- `config.js` — konfigurasi alamat API Google Apps Script.
- `Code.gs` — source backend untuk dipasang di project Google Apps Script; GitHub Pages tidak menjalankan file ini secara langsung.
- `.nojekyll` — menonaktifkan pemrosesan Jekyll pada GitHub Pages.

## Setup

1. Buka project Google Apps Script yang digunakan oleh aplikasi.
2. Pastikan backend `Code.gs` sudah menggunakan versi dari paket ini. Jika mengganti source, periksa kembali ID Spreadsheet, folder Drive, dan Script Properties yang dipakai project Anda sebelum deploy.
3. Deploy sebagai **Web app** dengan konfigurasi yang sesuai untuk aplikasi Anda. Setelah deploy, salin URL Web App yang berakhiran `/exec`.
4. Edit `config.js`, lalu ganti `PASTE_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE` dengan URL `/exec` tersebut. Jangan menambahkan tanda kutip atau karakter lain di luar string URL.
5. Upload `index.html`, `config.js`, `.nojekyll`, dan `README.md` ke root repository GitHub. `Code.gs` boleh disimpan di repository sebagai arsip/source, tetapi backend tetap dijalankan melalui Google Apps Script.
6. Aktifkan GitHub Pages di **Settings → Pages**, pilih branch dan folder root yang digunakan repository.
7. Tes URL Apps Script dengan menambahkan `?health=1` di belakang URL `/exec`. Endpoint seharusnya menampilkan JSON dengan status `online`.
8. Buka URL GitHub Pages dan tes alur panelis serta login admin.

## Catatan penting

- `config.js` masih memakai placeholder; GitHub Pages belum dapat tersambung sampai URL `/exec` yang benar diisi.
- Perubahan pada `Code.gs` tidak otomatis memperbarui backend yang sudah ter-deploy. Jika backend diubah, buat versi deployment baru melalui **Deploy → Manage deployments → Edit → New version → Deploy** agar URL deployment tetap sama.
- Jangan membagikan token admin atau kredensial Google di repository.
- Library eksternal (Tailwind, SweetAlert2, Chart.js, dan lainnya) dimuat melalui CDN sehingga browser membutuhkan koneksi internet.
