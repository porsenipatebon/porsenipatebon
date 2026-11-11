# PORSENI 2026 — PAC IPNU-IPPNU Patebon

## Struktur repository

- `index.html` — website publik.
- `admin.html` — Portal Panitia.
- `assets/images/` — logo dan ikon website.
- `gas/backend/` — source Google Apps Script (`Kode.js`) dan manifest.
- `.github/workflows/deploy-apps-script.yml` — workflow sinkronisasi/deploy Apps Script.
- `CNAME` — konfigurasi domain GitHub Pages.

## Catatan pemeliharaan

1. Jangan memindahkan `CNAME` dari root repository.
2. Jangan mengubah lokasi `gas/backend/` tanpa menyesuaikan `rootDir` pada konfigurasi clasp dan workflow.
3. Jika menambah aset gambar, simpan di `assets/images/` dan gunakan path tersebut dari HTML/CSS.
4. Apps Script dikelola melalui workflow manual pada GitHub Actions. Mode `push-only` hanya menyinkronkan source; `deploy-production` juga memperbarui deployment produksi.
5. Rahasia GitHub Actions (`CLASPRC_JSON`, `CLASP_JSON`, dan `GAS_DEPLOYMENT_ID`) tidak boleh dimasukkan ke repository.
6. `gas/backend/Kode.js` adalah sumber backend. `gas/admin.html` tidak diperlukan karena Portal Panitia aktif berada di root sebagai `admin.html`.
