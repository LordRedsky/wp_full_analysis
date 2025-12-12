# WP Deep Analysis – Code Guide

## Struktur Proyek
- `app.py` – UI Streamlit untuk menjalankan Exact Match dan Deep Analysis
- `exact_match.py` – logika pencocokan WP ↔ Sertifikat ↔ Kendali dan pengisian untuk cell putih
- `deep_analysis.py` – pemrosesan cell abu-abu → hijau khusus sesuai instruksi

## Utilitas Bersama (exact_match.py)
- Baca Excel: `exact_match.py:16–19`
- Normalisasi string/angka: `exact_match.py:22–76`
- Setter angka dan padding angka: `exact_match.py:77–88`
- Cari header kolom: `exact_match.py:90–96`
- Deteksi cell putih (unfilled): `exact_match.py:98–100`
- Keterangan bernomor: `exact_match.py:102–110`
- Deteksi gelar/awalan dan stripping: `exact_match.py:116–165`
- Warna: `exact_match.py:9–13`

## Exact Match – Pengisian Data (Referensi Cell Putih)
- Ambil kandidat dari Sertifikat yang cell `Nama` putih
- Jika cocok dan `NIB` ditemukan di Kendali:
  - `JENIS_LHN = JENIS_ZONA`
  - `KELAS_BUMI` → `_set_padded_number(..., 3)`
  - `NIB_1` → `_set_padded_number(..., 5)`
  - `KD_ZNT` salin apa adanya
  - `NJOP_BUMI` → `_set_number`
  - Validasi `LUAS_BUMI` vs `LUASTERTUL`; jika beda ganti, hijau + keterangan
  - Pewarnaan biru di Sertifikat + `KETERANGAN = "Data cocok"`
- Referensi: `exact_match.py:345–370`

## Deep Analysis – Abu-Abu → Hijau (deep_analysis.py)
- Hanya memproses `NAMA_WP` abu-abu, cocokkan ke `Nama` putih di Sertifikat
- Anti duplikasi `NIB_1` di WP: cek sebelum pilih kandidat (`deep_analysis.py:135–140`)
- Jika tidak ada kandidat putih / semua `NIB` terpakai → WP `NAMA_WP` merah + satu keterangan kasus
- Jika cocok:
  - `JENIS_LHN = JENIS_ZONA` (`deep_analysis.py:204`)
  - `KELAS_BUMI` dan `NIB_1` ditulis sebagai teks ber-padding tetap agar nilai tampilan dan klik sama (contoh: "00333"/"00654") (`deep_analysis.py:205–208`, helper `deep_analysis.py:43–56`)
  - `KD_ZNT` salin apa adanya (`deep_analysis.py:209`)
  - `NJOP_BUMI` angka via `_set_number` (`deep_analysis.py:210–214`)
  - `LUAS_BUMI` disesuaikan ke `LUASTERTUL` jika berbeda, hijau + keterangan (`deep_analysis.py:217–224`)
  - `NAMA_WP` diganti ke `Nama` asli Sertifikat, hijau + keterangan bernomor (`deep_analysis.py:226–231`)
  - Sertifikat: baris diwarnai biru + `Data cocok` (`deep_analysis.py:233–235`)

## Warna & Arti
- Putih: cell tanpa fill (sumber kandidat Sertifikat)
- Abu-abu: target Deep Analysis (WP)
- Merah: kegagalan (nama/NIB tidak ditemukan, atau NIB terpakai)
- Hijau: penyesuaian sukses (nama/luas)
- Biru: data di Sertifikat sudah cocok

## Aturan Keterangan
- WP, abu-abu → merah: keterangan lama dihapus, ganti satu keterangan baru
- WP, abu-abu → hijau: keterangan lama dihapus, ganti keterangan baru; jika >1, bernomor (1., 2., ...)
- Sertifikat: tetap penomoran otomatis saat menambah catatan (lihat `exact_match.py:102–110`)

## Penamaan File Output (app.py)
- Exact Match: `<nama file lama>_update.<ext>` (`app.py:38–41`, tombol unduh `app.py:71–83`)
- Deep Analysis: `<nama file lama>_deep analysis.<ext>` (`app.py:57–59`, tombol unduh `app.py:94–99`)
- Tim Pemetaan (salin WP deep tanpa baris merah, format & warna dipertahankan): `<nama file lama>_tim pemetaan.<ext>` (`app.py:58–59`, tombol unduh `app.py:100–107`)

## Invarian yang Harus Dijaga
- Deep Analysis hanya dijalankan setelah Exact Match
- Hanya memproses `NAMA_WP` abu-abu; kandidat Sertifikat harus putih
- Tidak ada `NIB_1` duplikat di WP setelah Deep Analysis
- Format kolom ber-padding hasil hijau di WP bertipe teks ber-padding tetap (contoh: "00333", "00654") via helper `_set_padded_text`
- Keterangan WP menimpa yang lama sesuai kasus, bernomor saat >1
- `exact_match.py` tidak diubah

## Jalankan Aplikasi
- Perintah: `streamlit run app.py`
- Urutkan: jalankan Exact Match → jalankan Deep Analysis → unduh hasil (WP deep updated, Tim Pemetaan, Sertifikat)
