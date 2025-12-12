# FINAL ANALYSIS

## PERHATIAN!!!
1. Final analysis dilakukan setelah semua analisis selesai dilakukan.
2. Proses utama Final analysis sama prosesnya dengan deep analysis,sehingga bisa menggunakan kode yang sama dengan deep analysis tetapi dibuat pada file yang berbeda, jangan gabung ke scrypt exact_match.py, deep_analysis.py dan analysis_sert.py. misal final_analysis.py.
3. file WP kosong yang digunakan adalah file WP yang sudah saya siapkan pada folder empty yaitu file WP NAMA DESA.xlsx. file ini nanti akan kita isi dengan data pada file sertifikat hasil deep analysis.
4. file sertifikat yang digunakan dalah file sertifikat hasil deep analysis jadi tidak perlu di upload ulang
5. file form kendali yang digunakan dalah file form kendali awal jadi tidak perlu di upload ulang
6. Buatkan scrypt tersendiri khusus untuk full analysisi ini, jangan gabung dengan scrypt exact_match.py, deep_analysis.py dan analysis_sert.py. misal final_analysis.py.

## TAHAP 03 - FINAL ANALYSIS
### Pengisian File WP Kosong
- buka file sertifikat dan copy semua data pada kolom NAMA yang hanya berwarna putih (no fill). selain itu jangan di copy. kemudian paste data tersebut ke file WP NAMA DESA.xlsx pada kolom NAMA_WP yang berwarna kuning.
- buka file sertifikat dan copy semua data pada kolom Luas yang hanya berwarna putih (no fill). selain itu jangan di copy. kemudian paste data tersebut ke file WP NAMA DESA.xlsx pada kolom LUAS_BUMI yang berwarna kuning.
- Pada file WP NAMA DESA.xlsx, isi kolom ALAMAT_OP yang berwarna kuning dengan nama desa. Nama desa di ambil dari nama file form kendali dengan cara mengambil bagian nama file form kendali setelah kata "Kendali". misal nama file form kendali adalah "FORM KENDALI BAHUTARA.xlsx", maka nama desa yang di isi adalah "DESA BAHUTARA". Jika nama filenya adalah "FORM KENDALI KATOBU.xlsx", maka nama desa yang di isi adalah "DESA KATOBU". Begitu seterusnya.
- setelah semua data terisi mulai dari NAMA_WP sampai LUAS_BUMI, maka file ini siap digunakan untuk melakukan tahapan full analysis

### Tahap Full Analysis
- File yang digunakan untuk analisis adalah file WP NAMA DESA.xlsx yang sudah terisi dengan data NAMA_WP, LUAS_BUMI dan ALAMAT_OP, file sertifikat hasil deep analysis dan File Form kendali.
- Cara kerja dan logic sama seperti exact_macth.py