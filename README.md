# Generator RSNI

Aplikasi Streamlit untuk mengoptimasi format, membuat cover, daftar isi,
prakata/pendahuluan, info pendukung, dan menerjemahkan dokumen `.docx`
standar (RSNI/ISO) secara otomatis.

## Menjalankan secara lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud (gratis)

1. Push repo ini ke GitHub (folder ini, jangan di-zip).
2. Buka [share.streamlit.io](https://share.streamlit.io), pilih repo ini,
   dan set **Main file path** ke `app.py`.
3. Deploy.

### Supaya aplikasi tidak "tidur" (sleep)

Streamlit Community Cloud otomatis menidurkan aplikasi yang tidak
menerima traffic selama 12 jam. Ini adalah kebijakan platform gratis
dan tidak bisa dimatikan dari dalam kode aplikasi — solusinya adalah
membuat traffic buatan secara berkala **sebelum** 12 jam tercapai.

Repo ini sudah menyertakan workflow GitHub Actions
(`.github/workflows/keep_alive.yml`) yang mengunjungi URL aplikasi
setiap 6 jam. Agar aktif:

1. Deploy aplikasi terlebih dahulu dan salin URL-nya
   (mis. `https://nama-app-anda.streamlit.app`).
2. Di GitHub, buka **Settings → Secrets and variables → Actions →
   New repository secret**.
3. Buat secret dengan nama `STREAMLIT_APP_URL` berisi URL tersebut.
4. Workflow akan berjalan otomatis tiap 6 jam (bisa juga dijalankan
   manual lewat tab **Actions → Keep Streamlit App Awake → Run workflow**).

Catatan: jika aplikasi sudah terlanjur tidur, ping biasa tidak cukup
untuk membangunkannya karena Streamlit menampilkan halaman "wake up"
yang butuh klik tombol. Selama workflow ini berjalan rutin sebelum
12 jam habis, aplikasi tidak akan sempat tidur sama sekali.

## Struktur file

- `app.py` — aplikasi Streamlit utama (UI + orkestrasi pipeline).
- `engine2.py` — optimasi format dasar dokumen.
- `engine4.py` — pembuatan cover.
- `engine5.py` — pembuatan daftar isi.
- `engine6.py` — prakata & pendahuluan.
- `engine7.py` — info pendukung.
- `engine9.py` — mesin terjemahan & kamus istilah.
- `engine8.py` — modul cadangan (belum dipakai oleh `app.py`).
