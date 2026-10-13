============================================================
  PENCocokan Transaksi v4.3.2
============================================================

APLIKASI INI UNTUK APA?
-----------------------
Mencocokkan transaksi Faktur Pembelian dengan Penerimaan
Barang secara otomatis. Cocok untuk rekonsiliasi supplier.

CARA PAKAI
----------
1. Klik dua kali "PencocokanTransaksi.exe"
2. Tab "Data" -> Klik "Import Excel ke Database"
3. Pilih file Excel, pilih sheet, atur pemetaan kolom
4. Klik "Import Sekarang"
5. Tab "Pencocokan":
   - Supplier: pilih supplier (atau ALL)
   - Maks baris per grup: 30 atau 50
   - Time budget: 2000 ms
6. Klik "Jalankan Pencocokan"
7. Tab "Hasil" -> lihat dan ekspor

SYARAT SISTEM
-------------
- Windows 10/11 64-bit
- RAM minimal 4 GB (8 GB direkomendasikan)
- Tidak perlu install Python, Excel, atau apapun
- File Excel yang diimport format .xlsx atau .xlsm

CATATAN PENTING
---------------
- Aplikasi ini menyimpan data di file "data.db" di folder
  yang sama dengan .exe. JANGAN hapus file itu kalau ingin
  data tetap tersimpan.
- Untuk supplier dengan ratusan baris, proses Deep Search
  bisa memakan waktu 2-5 menit. Jangan tutup aplikasi.
- Tombol "Cancel" untuk menghentikan proses kapan saja.

FILE YANG DIHASILKAN
--------------------
- data.db (database lokal, otomatis dibuat)
- data.db-wal, data.db-shm (SQLite temp, bisa diabaikan)
- hasil_*.csv / *.xlsx (hasil ekspor)

TROUBLESHOOTING
---------------
Q: Aplikasi tidak mau terbuka / hilang setelah klik?
A: Klik kanan .exe -> Properties -> Unblock.
   Kalau masih tidak jalan, cek Windows Defender.

Q: Hasil pencocokan masih banyak yang merah?
A: Coba:
   - Naikkan "Time budget" ke 5000 ms
   - Naikkan "Maks baris per grup" ke 50 atau 100
   - Jalankan per supplier (jangan ALL) untuk hasil lebih baik

Q: Butuh kirim file Excel ke orang lain?
A: Bisa pakai fitur "Ekspor Excel Bertanda" di tab Hasil.

============================================================