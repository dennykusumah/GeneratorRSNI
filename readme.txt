GENERATOR RSNI — ISO TO RSNI CONVERTER
======================================

Generator RSNI adalah aplikasi Streamlit untuk memformat dan menerjemahkan
dokumen standar ISO menjadi draft RSNI melalui pipeline Engine 1 sampai
Engine 9.


1. PERSYARATAN SISTEM
---------------------

- Windows 10 atau Windows 11 direkomendasikan.
- Python 3.10 atau versi yang lebih baru.
- Microsoft Word desktop wajib tersedia untuk fungsi akhir Engine 9.
- Koneksi internet diperlukan untuk:
  - Glosarium SNI di Google Sheets;
  - Glosarium Istilah Asing di Google Sheets; dan
  - layanan terjemahan yang digunakan Engine 8.

Catatan penting:
Engine 9 menggunakan Microsoft Word melalui pywin32 untuk menghitung halaman,
mendeteksi halaman kosong setelah TOC, menghapus halaman kosong tersebut, dan
menjalankan "Update page numbers only". Fitur ini tidak dapat dijalankan pada
Linux, Streamlit Community Cloud, atau komputer Windows tanpa Microsoft Word.


2. SUSUNAN DAN NAMA FILE
------------------------

Letakkan seluruh file berikut dalam satu folder:

app.py
engine1.py
engine2.py
engine3.py
engine4.py
engine5.py
engine6.py
engine7.py
engine8.py
engine9.py
requirements.txt
readme.txt

Nama file harus persis seperti daftar di atas. Hapus tambahan tanggal, nomor,
dan tanda kurung dari nama file hasil unduhan. Contoh:

app(20260825-120619).py       menjadi app.py
engine1(20260825-120619).py   menjadi engine1.py
engine2(7).py                 menjadi engine2.py
engine3(6).py                 menjadi engine3.py
engine4(6).py                 menjadi engine4.py
engine5(10).py                menjadi engine5.py
engine6(9).py                 menjadi engine6.py
engine7(20260825-120619).py   menjadi engine7.py
engine8(20260825-120618).py   menjadi engine8.py
engine9(20260825-120618).py   menjadi engine9.py


3. INSTALASI DI WINDOWS
-----------------------

Buka PowerShell pada folder aplikasi, lalu jalankan:

py -m venv venv
venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt

Jika PowerShell menolak aktivasi virtual environment, jalankan perintah berikut
untuk sesi PowerShell tersebut, kemudian ulangi aktivasi:

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
venv\Scripts\Activate.ps1


4. MENJALANKAN APLIKASI
-----------------------

Pastikan virtual environment aktif, kemudian jalankan:

streamlit run app.py

Browser biasanya terbuka otomatis. Jika tidak, buka alamat lokal yang tampil
pada PowerShell, umumnya:

http://localhost:8501


5. CARA MENGGUNAKAN
-------------------

1. Unggah dokumen ISO berformat .docx.
2. Isi nomor SNI dan nomor ICS.
3. Klik tombol "Proses".
4. Tunggu progres Engine 1 sampai Engine 9 mencapai 100%.
5. Unduh dokumen hasil Engine 9.

Jangan menutup Microsoft Word secara paksa ketika Engine 9 sedang berjalan.
Engine 9 membuka Word secara tersembunyi untuk memperbarui TOC dan pagination.


6. FUNGSI PIPELINE
------------------

- Engine 1: menyiapkan dan memangkas bagian Introduction/Content.
- Engine 2: membuat halaman cover dan metadata dokumen.
- Engine 3: menyiapkan halaman Daftar Isi.
- Engine 4: membuat Prakata dan Pendahuluan.
- Engine 5: menduplikasi serta menandai area konten yang diperlukan.
- Engine 6: menambahkan informasi pendukung dokumen SNI.
- Engine 7: menerapkan format Daftar Isi, konten, dan Bibliografi.
- Engine 8: menerjemahkan zona yang ditentukan dengan Glosarium SNI dan
  Glosarium Istilah Asing, sekaligus mempertahankan format dokumen.
- Engine 9: menerapkan style final, membuat TOC, menghapus halaman kosong
  setelah TOC, dan memperbarui nomor halaman TOC melalui Microsoft Word.


7. SUMBER GLOSARIUM
-------------------

Glosarium SNI:
https://docs.google.com/spreadsheets/d/1BBPCMPwvbBk5LPdoDQwnjQzcPHv7_RDKENqeMsklF-8/edit?usp=sharing

Glosarium Istilah Asing:
https://docs.google.com/spreadsheets/d/1NZm1HjsjxmflxnZlzV_O2XF75ZlMUOu8VVofsKfp_FA/edit?usp=sharing


8. PENYELESAIAN MASALAH
-----------------------

A. ModuleNotFoundError
   Pastikan virtual environment aktif, lalu jalankan kembali:

   pip install -r requirements.txt

B. Error "cannot import name" dari engine
   Pastikan nama file sudah menjadi engine1.py sampai engine9.py dan tidak ada
   file engine versi lama dengan nama yang sama di folder aplikasi.

C. Error deep_translator
   Jalankan:

   pip install --upgrade deep-translator

D. Error pythoncom atau win32com
   Jalankan pada Windows:

   pip install --upgrade pywin32

   Tutup dan buka kembali PowerShell setelah instalasi bila modul masih belum
   terbaca.

E. Microsoft Word tidak ditemukan
   Pastikan Microsoft Word desktop sudah terpasang dan dapat dibuka normal.
   Word versi web tidak menyediakan otomatisasi COM yang diperlukan Engine 9.

F. Glosarium atau terjemahan tidak dapat dimuat
   Periksa koneksi internet serta pastikan kedua Google Sheets masih dapat
   diakses oleh aplikasi.

G. Port Streamlit sedang digunakan
   Jalankan aplikasi pada port lain:

   streamlit run app.py --server.port 8502


9. CATATAN OPERASIONAL
----------------------

- Gunakan salinan dokumen sumber agar file asli tetap aman.
- Hindari menjalankan dua proses terhadap file yang sama secara bersamaan.
- File sementara aplikasi dibersihkan otomatis oleh app.py.
- Periksa hasil akhir, TOC, penomoran, format italic, superscript, dan
  subscript sebelum dokumen digunakan sebagai draft resmi.


© 2026 Generator RSNI — ISO to RSNI Converter
Developed by Denny Kusuma H.
