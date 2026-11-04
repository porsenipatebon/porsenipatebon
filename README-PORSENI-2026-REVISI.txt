PAKET REVISI WEBSITE PORSENI 2026
PAC IPNU-IPPNU PATEBON

ISI PAKET
- index.html                       Website utama PORSENI 2026
- admin.html                       Portal Panitia dengan kontrol data dan hasil kejuaraan
- Code_GS_PERBAIKAN_BYTE_GAMBAR.gs Backend Google Apps Script
- logo-porseni.png                 Logo resmi
- kop-porseni-resmi.png            Kop resmi
- favicon-porseni.png              Favicon PNG
- favicon.ico                      Favicon ICO
- porseni-o-mark.png               Crop huruf O warna-warni dari kop resmi untuk wordmark header

PENEMPATAN FILE WEBSITE
Unggah index.html, admin.html, logo-porseni.png, kop-porseni-resmi.png, favicon-porseni.png,
favicon.ico, dan porseni-o-mark.png ke direktori hosting yang sama. Jangan mengubah nama file
porseni-o-mark.png karena dipanggil langsung oleh index.html.

PEMBARUAN BACKEND (WAJIB AGAR FITUR ADMIN BARU AKTIF)
1. Buka project Google Apps Script yang terhubung dengan spreadsheet database PORSENI.
2. Cadangkan source lama, lalu perbarui file script menggunakan isi Code_GS_PERBAIKAN_BYTE_GAMBAR.gs.
3. Pastikan Script Property ADMIN_PASSWORD tetap terisi dengan password admin bersama yang berlaku.
4. Simpan, kemudian Deploy > Manage deployments > Edit deployment Web App > pilih New version > Deploy.
5. Pertahankan konfigurasi Web App yang digunakan website sekarang: Execute as akun pemilik script dan akses sesuai konfigurasi publik form yang sudah berjalan.
6. URL endpoint di index.html dan admin.html tetap menggunakan URL Web App yang saat ini sudah tercantum di file.

IDENTITAS LOGIN ADMIN
Dropdown login berisi Tyo, Sochibul, dan Fuad. Ketiganya menggunakan password admin bersama yang tersimpan di Script Properties.
Nama petugas disimpan pada sesi dan dicatat di kolom admin pada Log Admin.

FITUR PORTAL ADMIN
- Dashboard pendaftaran kontingen dan berkas cabang lomba.
- Edit data pendaftaran yang diizinkan.
- Ganti atau tambah lampiran pada folder Drive pendaftaran.
- Hapus baris data kontingen atau cabang lomba. File/folder Drive sengaja tetap dipertahankan sebagai arsip.
- Verifikasi, catatan panitia, pengiriman email, dan pembukaan WhatsApp.
- Ekspor daftar pendaftaran ke CSV.
- Input, edit, dan hapus juara 1, 2, dan 3 per cabang/kategori.
- Perhitungan medali dan klasemen juara umum otomatis.
- Semua perubahan penting dicatat pada sheet Log Admin.

HASIL KEJUARAAN WEBSITE UTAMA
Saat hasil pertama kali disimpan dari Portal Panitia, backend otomatis membuat sheet
"Hasil Kejuaraan". Website utama membaca daftar juara dan rekap medali dari endpoint public_recap.

STATUS SEMENTARA / FINAL DAN PDF RESMI
- Website menampilkan label HASIL SEMENTARA sampai petugas menekan "Finalisasi Juara Umum" di menu Hasil Kejuaraan pada Portal Panitia.
- Finalisasi mensyaratkan Juara 1 sudah diisi pada seluruh 24 kategori. Jika belum lengkap, backend menolak finalisasi dan menyebut kategori yang masih kosong.
- Setelah finalisasi, website menampilkan status HASIL RESMI dan tombol "Cetak PDF Hasil Resmi PORSENI 2026".
- PDF berisi Juara Umum PORSENI 2026, Rekap Perolehan Juara, dan daftar hasil seluruh cabang/kategori lomba, termasuk kategori yang belum memiliki Juara 2 atau Juara 3. Kop memakai aset kop-porseni-resmi.png yang sama dengan formulir cetak.
- Setiap perubahan atau penghapusan hasil setelah finalisasi otomatis membatalkan status final. Admin perlu memeriksa dan melakukan finalisasi kembali.
- Website menampilkan seluruh daftar cabang dan kategori sejak awal, walaupun hasil belum diinput. Urutan mengikuti katalog lomba pada backend.
- Rekap homepage menyimpan cache data publik terakhir di browser. Jika koneksi sementara gagal, data terakhir tetap dapat ditampilkan dengan label bahwa data belum diperbarui.

CATATAN PENTING
- Sintaks JavaScript pada file HTML dan Apps Script telah diperiksa secara lokal.
- Integrasi langsung dengan Google Sheets/Drive belum dapat diuji dari lingkungan lokal ini.
  Karena itu, deploy ulang Apps Script diperlukan sebelum fitur backend dipakai di website live.
- Penggantian lampiran membuat file baru di folder Drive dan memperbarui tautan di spreadsheet.
  File lama tidak dihapus otomatis agar arsip tidak hilang.
- Penghapusan data dari portal menghapus baris pada sheet terkait, bukan folder Drive.
