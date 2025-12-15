kamu adalah seorang yang ahli dalam analisis data dan programing menggunakan python dengan pengalaman 5 tahunan, jadi kamu sudah sangat ekpert. Saya akan share projek yang sedang saya buat Dimana saya akan menganalisa 3 file excell yang berisi ribuan raw data yang akan saya cocokkan dan saya update. Klasifikasi file dan proses berpikir saya jelaskan dibawah ini. Tolong kerjakan dengan teliti

baik, saya mempunyai 3 buah file yang akan saya Analisa dan cocokkan satu sama lain.

*File pertama yaitu file nama wajib pajak yang berisi :*\
`NAMA_WP` (sudah ada isian data);\
`JENIS_LHN` (kosong yang nanti akan di isi);\
`KELAS_BUMI` (kosong yang nanti akan di isi);\
`NIB_1` (kosong yang nanti akan di isi);\
`KD_ZNT` (kosong yang nanti akan di isi);\
`NJOP_BUMI` (kosong yang nanti akan di isi);\
`LUAS_BUMI` (sudah ada isian yang nanti akan di cocokkan);\
\
*File kedua yaitu file sertifikat yang berisi :*\
`Nama` (sudah ada isian data);\
`NIB` (sudah ada isian data);\
`Luas` (sudah da isian data);\
\
*File ketiga yaitu file form kendali yang berisikan semua data yang dibutuhkan untuk menjadi bahan acuan mengisi cell kosong pada file pertama. file ini berisikan :*\
`NIB` (sudah ada isian data)\
`LUASTERTUL` (sudah ada isian data)\
`JENIS_ZONA` (sudah ada isian data)\
`KELAS_BUMI` (sudah ada isian data)\
`KD_ZNT` (sudah ada isian data)\
`RPBULAT` (sudah ada isian data)

FLOW PENGERJAAN \
***UPLOAD FILE*** \
upload ke 3 file tersebut yaitu file wajib pajak, file sertifikat dan file form kendali dalam bentuk excell. \
***PENYESUAIAN KOLOM PADA FILE SEBELUM DI ANALISA*** \
penyesuaian kolom pada file sertifikat
- ubah nilai kolom NIB menjadi angka (convert to number, format number 123, bukan 123.0)
- ubah nilai kolom Luas menjadi angka (convert to number, format number 123, bukan 123.0)
- mengidentifikasi nama para ahli waris pada kolom Nama dengan cara sebagai berikut :
- screening semua cell yang ada pada kolom `Nama` dari data pertama sampai terakhir dan pastikan pada masing-masing cell hanya terdiri dari satu nilai (nama). Hal ini kita lakukan dengan cara mengecek jumlak karakter dan jumlah nama pada cell tersebut. Jika panjang karakter tidak melebihi 50 karakter maka pada cell tersebut hanya terdiri dari satu nama saja dan jika nama tersebut tidak mengandung koma (,) setelah nama atau gelar. 
Contoh nama dengan gelar `“LA ODE ABDUL GANJAL, S.SOS”`, data ini hanya terdiri dari satu nama saja, karena setelah gelar S.SOS tidak diikuti dengan tanda koma atau nama dengan tanpa gelar `“LA ODE ABDUL GANJAL”` juga hanya terdiri dari satu nama, karena setelah nama tidak diikuti dengan randa koma (,). jika nilai pada kolom `Nama` tersebut lebih dari 50 karakter atau diakhir nama atau gelar terdapat tanda (,) maka bisa dipastikan bahwa cell tersebut berisi data nama ahli waris. Contoh nama dengan data ahli waris : `"LA ODE ABDUL GANJAL, S.SOS, LA ODE ABDUL MUHIDDIN, LA ODE MUH. UNTUNG, LA ODE MUSABAKA, WA LITA, WA ODE SITTI HADIAH, WA ODE SITTI MAULID"`. Bisa dilihat bahwa pada data ini panjang karakter melebihi 50 karakter dan ada beberapa nama yang dipisahkan oleh koma. ketika ketemu data seperti ini (yang terdiri lebih dari satu nama) langsung beri warna kuning pada row tersebut, sebagai tanda bahwa data tersebut dalam format ahli waris dan beri keterangan `"Data ahli waris"`.

penyesuaian kolom pada file form kendali
-	ubah nilai kolom NIB menjadi angka (convert to number, format number 123, bukan 123.0)
<!-- -	ubah nilai kolom LUASTERTUL menjadi angka (convert to number, format number 123, bukan 123.0) -->
-	ubah nilai pada kolom `RPBULAT` menjadi angka (convert to number) dengan menghilangkan Rp, koma dan titik.

penyesuaian kolom pada file wajib pajak
<!-- -	ubah nilai kolom `LUAS_BUMI` menjadi angka (convert to number, format number 123, bukan 123.0) -->

### TAHAP 1 - EXACT MATCH

***AMBIL DATA DARI FILE WAJIB PAJAK***\
- ambil data pada kolom `NAMA_WP`.\
<!-- ***CARI DATA `NAMA_WP` YANG SAMA DENGAN `Nama` DI FILE SERTIFIKAT*** -->
- Cocokkan `NAMA_WP` pada file wajib pajak dengan `Nama` pada file sertifikat (NAMA_WP ===  NAMA). Jika cell sudah berwarna biru, kuning, atau merah → cari alternatif cell dengan nama sama tapi warna cell masih putih.
- Jika cell berwarna putih dan data cocok → ambil `NIB`.
- Jika tidak ada yang cocok pada cell berwarna putih, lakukan pra-pemeriksaan nama berikut:
  - Cek gelar pada `NAMA_WP` (di awal/akhir nama). Gelar umum yang diakui: DR, DR., DRS, DRS., IR, IR., H, H., HJ, HJ., KH, KH., S.SOS, S.Sos, S.Ag, S.Pd, S.H, S.E, S.Kom, S.IP, S.Psi, S.Kep, S.T, ST, M.Si, M.Sos, M.Kom, M.Pd, M.H, M.E, M.Ak, MT, MSc, MA, PhD, Sp., SpA, SpB, SpOG, beserta padanan titik/kapitalnya.
  - Cek nama depan/prefix keluarga pada `NAMA_WP`. Contoh umum: LA ODE, LAODE, WA ODE, WAODE, MUH, M, LD, WD, LA, WA.
  - Jika `NAMA_WP` mengandung salah satu dari gelar atau prefix di atas, isi `KETERANGAN = "Nama Wajib Pajak tidak ditemukan di sertifikat"` pada file Wajib Pajak dan ubah warna cell `NAMA_WP` menjadi abu-abu.
***MENCOCOKKAN NILAI `NIB` FILE SERTIFIKAT DAN `NIB` FORM KENDALI***
- Cari `NIB` di kolom NIB pada Form Kendali.
- Jika tidak ada → isi `KETERANGAN = "NIB tidak ditemukan pada form kendali"` pada file sertifikat dan file Wajib Pajak, lalu beri warna merah hanya pada kolom `NAMA_WP` pada cell tersebut pada file Wajib Pajak dan kolom `NAMA` pada row tersebut pada file sertifikat.
- Jika datanya cocok (`NIB` pada file sertifikat === `NIB` pada file form kendali), maka lanjutkan ke tahap berikutnya yaitu pengambilan data kendali
- Ambil data: `NIB`, `RPBULAT`, `JENIS_ZONA`, `KELAS_BUMI`, `KD_ZNT`, `LUASTERTUL`.
- Lakukan validasi luas dengan cara bandingkan `LUASTERTUL` form kendali vs `Luas` dari Sertifikat:
Jika berbeda → gunakan **nilai dari Sertifikat**.
Jika sama → tetap gunakan `LUASTERTUL`.\
***UPDATE FILE WAJIB PAJAK DENGAN DATA KENDALI YANG SUDAH KITA AMBIL***
- Lakukan update pada data wajib pajak dengan mengisi kolom-kolom berikut dengan data yang sudah kita punya.
- `JENIS_LHN = JENIS_ZONA`
- `KELAS_BUMI = KELAS_BUMI (dari form kendali)`
- `NIB_1 = NIB (awal pada file sertifikat)`
- `KD_ZNT = KD_ZNT (dari form kendali)`
- `NJOP_BUMI = angka dari RPBULAT (hilangkan Rp, koma, titik jika masih ada)`
- sebelum mengisi `LUAS_BUMI`, cocokkan terlebih dahulu data pada `LUAS_BUMI` dengan `LUASTERTUL`.
  Jika berbeda → ganti dengan `LUASTERTUL` dan beri keterangan `"Telah dilakukan penyesuaian luas bumi sesuai data sertipikat. <LUAS_BUMI> -> <LUASTERTUL>"` pada kolom `KETERANGAN` pada file wajib pajak, dan warnai hijau hanya pada cell `LUAS_BUMI`.
  Jika nilainya sama → tetap gunakan nilai pada `LUAS_BUMI`.\
- kemudian Warnai row (`Nama`, `NIB`, `Luas`) dengan biru kemudian -> isi `KETERANGAN = "Data cocok"` pada file sertifikat.

### **Special Case Pada Keterangan**
-Jika terdapat dua keterangan atau lebih pada satu NIB, maka tolong pada kolom keterangan diberi penomoran. 1,2,3 dan seterusnya. biar saya bisa tau telah terjadi update apa saja pada data tersebut

buatkan scrypt python untuk proses exact match ini pada file exact match.py dan jalankan scrypt tersebut pada app.py agar mudah diimplementasikan.

## 🛠️ Implementasi Teknis
Kamu bisa implementasi dengan **Python + Pandas + OpenPyXL**:
- **Pandas** → untuk join/matching data antar file.
- **OpenPyXL** → untuk manipulasi warna cell di Excel.


*RUN*
streamlit run app.py