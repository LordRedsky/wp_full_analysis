# Analysis Sertificate

## PERHATIAN!!!
1. Analysis sertificate dilakukan pertamakali sebelum analisis yang lain dilakukan
2. Analysis sertificate hanya dilakukan pada file Sertifikat pada kolom "NIB" hanya pada data yang berwarna putih (no fill)
3. Buatkan script tersendiri khusus untuk Analysis Sertificate jangan gabung ke scrypt exact_match.py dan deep_analysis.py, misalnya analysis_sert.py
4. Script analysis_sert.py harus memiliki fitur untuk memeriksa apakah data pada kolom "NIB" memiliki nilai atau tidak dan nilainya itu ada duplikate atau tidak.

### TAHAP 00 - ANALYSIS SERTIFICATE
- ambil data pada kolom "NIB" yang hanya berwarna putih atau no fill
- periksa apakah data pada kolom "NIB" memiliki nilai atau tidak
  - jika tidak ada atau kosong, maka beri keterangan "NIB kosong" dan beri warna merah pada row tersebut (NAMA, NIB dan Luas)
- periksa apakah data pada kolom "NIB" memiliki duplikate atau tidak
  - jika ada duplikate, maka beri keterangan "NIB memiliki duplikate" dan beri warna merah pada row tersebut (NAMA, NIB dan Luas)
  - contoh kasus
    - data pada kolom "NIB" memiliki nilai "1234567890"
    - data pada kolom "NIB" memiliki nilai "0987654321"
    - data pada kolom "NIB" memiliki nilai "1234567890" (duplikate). nah data duplikate ini yang diwarnai merah. data awalnya tidak diwarnai merah karena dia adalah data pertama yang ditemui dengan nilai "1234567890".
- Setelah selesai menganalisa semua data NIB, simpan file Sertifikat dengan nama "Sertifikat_analysed.xlsx" dan file ini nanti akan digunakan sebagai file acuan untuk melakukan analisis lebih lanjut yaitu pada exact_match.py dan deep_analysis.py.

