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

### Deep Analysis Sertificate for Ahli Waris
- ambil data pada kolom "NAMA" yang hanya berwarna putih atau no fill
- Periksa apakah nama pada sertifikat memiliki data ahli waris atau tidak dengan mengecek apakah terdapat lebih dari satu nama induk pada cell tersebut
  - Buang gelar dan nama depan, ambil nama induknya. jika terdapat dua nama atau lebih, maka itu adalah ahli waris. Pada kasus kita tadi terdapat 4 nama induk, maka beri keternagn "Data Ahli Waris" dan beri warna kuning pada row tersebut (Nama, NIB dan Luas)
- contoh : nilai kolom NAMA "HAFIDIN, A.MA.PD.I, Drs. Amin, Wa Uceng, La Deli" disini terdapat empat nama yaitu Hafidin, Amin, Uceng dan Deli.
  Hafidin, A.MA.PD.I memiliki nama induk Hafidin, dengan membuang gelarnya yaitu A.MA.PD.I dan Drs. Amin memiliki nama induk Amin, dengan membuang gelarnya yaitu Drs.. dan seterusnya.
  
