# Validasi Export Modul

Aplikasi untuk membandingkan **file input awal** dengan **file export module**, lalu membuat file **Hasil** (.xlsx) untuk tiap pasangan. Kolom `check` di tiap Hasil berisi formula `=IF(EXACT(...),"Sama","Beda")`.

Aplikasi bisa dipakai dengan dua cara:
- **Lokal:** berjalan di laptop kamu, dibuka di browser. Tidak ada data yang dikirim ke internet.
- **Online:** di-deploy ke Streamlit Community Cloud, lalu dibuka lewat link. File yang diupload diproses di server Streamlit.

## Menjalankan lokal

**Windows (paling mudah):** double-click `jalankan.bat`. Saat pertama kali dijalankan, file ini membuat environment dan menginstall library otomatis. Setelah itu browser terbuka di `http://localhost:8501`. Jangan tutup jendela hitamnya selama aplikasi dipakai.

**Manual (Windows, Mac, atau Linux):**

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate    Mac/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Python 3.10 atau lebih baru diperlukan.

## Deploy online

1. Upload isi project ke repo GitHub (lihat daftar file di bawah).
2. Buka share.streamlit.io, login dengan GitHub, lalu pilih **Create app**.
3. Isi repo, branch `main`, dan Main file path `app.py`. Pilih Python 3.11 atau lebih baru.
4. Deploy, lalu atur siapa yang boleh membuka lewat menu Share atau Settings > Sharing.

Yang **jangan** diupload ke GitHub: folder `.venv`, `hasil`, `contoh`, `__pycache__`, dan semua file `.xlsx`.

## Alur pemakaian

**1. Mapping**
Upload satu atau banyak file mapping (format sama dengan sheet **Mapping**):

| nama file | nama module | column file | column module |
|---|---|---|---|
| 01.A.1_Rincian_Aset_Tetap_per_Fungsi_DM(1) | INP.PPE.030 Aset Tetap Fungsi | template_code | template_code |

Semua file mapping digabung jadi satu. Kolom `column module` yang kosong tidak dibandingkan (misalnya `file_path`). Klik **Pakai mapping ini**. Mapping berlaku selama halaman dibuka. Kalau halaman di-refresh, upload ulang.

**2. Pengaturan kunci**
Untuk tiap module, pilih satu atau lebih kolom sebagai **kunci**. Baris input dan export dipasangkan berdasarkan nilai kunci, bukan urutan baris. Kalau kunci dikosongkan, pasangan dibuat per posisi baris.

**3. Proses validasi**
- Upload file input awal, boleh banyak, format `.xlsx` atau `.csv`.
- Upload file export module, boleh banyak, format `.xlsx` atau `.csv`.
- Aplikasi membuat tabel pasangan. File input dicocokkan dengan "nama file" di mapping, lalu file export dicari lewat nama module. Kalau ada yang salah, ubah langsung di tabel. File input yang belum punya pasangan akan dilewati.
- Klik **Proses semua**. Tiap pasangan menghasilkan satu file `Hasil_<nama file input>.xlsx`.
- Download semua hasil sekaligus lewat **zip**, atau satu-satu lewat menu detail.

## Mencoba dengan data contoh

```bash
python buat_contoh.py
```

Ini membuat folder `contoh/` berisi:
- `mapping_contoh.xlsx`, `1.A.1 Aset tetap.xlsx`, `INP.ANR.AT.xlsx` (satu pasangan)
- `batch/` berisi tiga pasangan: input `.xlsx` dan export `.csv` (satu dengan pemisah `;` dan encoding Windows)

## Catatan penting

- **Perbandingan mengikuti EXACT Excel.** Huruf besar/kecil dan spasi dihitung. Angka `1000` dan teks `"1000"` dianggap **Sama**. Tanggal dibandingkan sebagai nomor serialnya, jadi tanggal yang tampil beda format bisa terlihat **Beda**.
- **Isi CSV dibaca sebagai teks.** Tanggal di CSV (misalnya `2025-01-01`) dibandingkan dengan nomor serial tanggal di Excel akan terlihat **Beda**. Pemisah (`,` atau `;`) dan encoding (UTF-8 atau Windows-1252) dideteksi otomatis.
- **Angka Beda di aplikasi adalah perkiraan.** Sumber kebenarannya adalah kolom `check` di file Hasil.
- Baris export yang tidak punya pasangan di input ditaruh di bagian bawah Hasil, dengan sisi input kosong, sehingga kolom check-nya **Beda**.
- Kalau satu nilai kunci muncul lebih dari sekali, pasangannya diurutkan (kemunculan ke-1 dengan ke-1, dan seterusnya). Aplikasi akan memberi peringatan.
- Satu file export tidak boleh dipasangkan ke dua file input. Aplikasi akan menolak proses kalau itu terjadi.
- Nama output dibuat unik. Kalau dua file input punya nama dasar yang sama, output kedua diberi akhiran `_2`.
- Mode online: mapping, pengaturan kunci, dan hasil tidak disimpan di server. Semuanya hidup di sesi browser masing-masing.

## Menjalankan tes (opsional)

```bash
pip install pytest
python -m pytest -q
```

## Struktur folder

```
app.py              antarmuka Streamlit (3 menu)
jalankan.bat        shortcut untuk Windows
requirements.txt    daftar library
.streamlit/         pengaturan Streamlit (tanpa telemetri)
inti/
  baca_excel.py     baca xlsx dan csv, header di baris 1
  mapping.py        baca dan gabung file mapping
  pencocokan.py     pasangkan baris (kunci / posisi) dan peniruan EXACT
  proses.py         cek kolom lalu hasilkan Hasil untuk satu pasangan
  batch.py          proses banyak pasangan, tebak pasangan, buat zip
  tulis_hasil.py    tulis file .xlsx dengan formula check
buat_contoh.py      membuat data contoh
tests/              tes otomatis
```
