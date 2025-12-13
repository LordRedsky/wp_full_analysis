# Deep Analysis

## PERHATIAN!!!
1. Deep Analysis dilakukan setelah proses exact_match.py.
2. Deep Analysis hanya menganalisa kolom `NAMA_WP` yang berwarna abu-abu saja. selain warna abu-abu, jangan disentuh
3. Buatkan script tersendiri khusus untuk deep analysis jangan gabung ke scrypt exact_match.py, misalnya deep_analysis.py
4. script deep_analysis.py akan melakukan penyesuaian nama pada kolom `NAMA_WP` yang berwarna abu-abu.
5. script deep_analysis.py akan mengubah nilai `NAMA_WP` menjadi nilai `Nama` yang sudah di penyesuaian nama depan dan gelar.
6. script deep_analysis.py akan mengubah nilai `KETERANGAN` menjadi "Telah dilakukan penyesuaian nama sesuai sertifikat. <`NAMA_WP`> -> <`Nama`>"
7. script deep_analysis.py hanya dijalankan ketika script exact_match.py selesai.

### TAHAP 2 - DEEP ANALYSIS
- ambil data pada kolom `NAMA_WP` yang hanya berwarna abu-abu. Jika setelah dicek semua data pada kolom `NAMA_WP` dan tidak ditemukan cell yang berwarna abu-abu, maka deep analysis tidak perlu dilanjutkan lagi.
- Jika ada cell berwarna abu-abu maka lakukan penyesuaian terlebih dahulu pada data tersebut sebelum lanjut ke step berikutnya dengan cara sebagai berikut :
  - cek `NAMA_WP` apakah memiliki gelar atau nama depan.
  - jika memiliki gelar dan nama depan, hilangkan gelar dan nama depannya terlebih dahulu, karena data yang digunakan untuk pencocokan nantinya adalah nama induk tanpa nama depan dan gelar.
  - jika nama tersebut tidak mengandung nama depan dan gelar, maka nama tersebut langsung bisa dipakai tanpa penyesuaian terlebih dahulu \
- kemudian kita mulai mencocokkan dengan data pada kolom `Nama` yang ada pada file sertifikat yang cell nya berwarna putih. selain warna putih, di skip saja. intinya cari cell yang hanya berwarna putih.
- Setelah ketemu cell warna putih maka langkah selanjutnya dalah melakukan pengondisian data `Nama` sebagai berikut :
  - cek cell `Nama` apakah memiliki gelar atau nama depan.
  - jika memiliki gelar dan nama depan, hilangkan gelar dan nama depannya terlebih dahulu, karena data yang digunakan untuk pencocokan nantinya adalah nama induk tanpa nama depan dan gelar.
  - jika nama tersebut tidak mengandung nama depan dan gelar, maka nama tersebut langsung bisa dipakai tanpa penyesuaian terlebih dahulu \
- kemudian cocokkan `NAMA_WP` yang sudah di kondisikan dan `Nama` yang sudah dikondisikan juga nanti. Pastikan semua cocok by karakter dan panjang karakter. Contoh nama pada `NAMA_WP` dan `Nama` yang cocok BY KARAKTER DAN PANJANG KARAKTER adalah "AMIN" === "AMIN" atau "NUR FIRA" === "NUR FIRA". Sedangkan data yang tidak cocok adalah "AMIR" !== "AMIRUDDIN" atau "NUR FIRA" !== "NUR FIRAS". Ingat, pastikan semua cocok BY KARAKTER DAN PANJANG KARAKTER.
- Jika cocok maka ambil nilai `NIB`
- Cari `NIB` di kolom NIB pada Form Kendali.
- Jika tidak ada → isi `KETERANGAN = "NIB tidak ditemukan pada form kendali"` pada file sertifikat dan file Wajib Pajak, lalu beri warna merah hanya pada kolom `NAMA_WP` pada cell tersebut pada file Wajib Pajak dan kolom `Nama` pada row tersebut pada file sertifikat.
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
- Ganti nilai `NAMA_WP` dengan nilai awal `Nama` pada sertifikat sebelum dilakukan penyesuaian nama depan dan gelar lalu isi keterangan "Telah dilakukan penyesuaian nama sesuai sertifikat. <`NAMA_WP`> -> <`Nama`>" lalu beri ganti warna abu-abu menjadi hijau pada cell tersebut.
- kemudian Warnai row (`Nama`, `NIB`, `Luas`) dengan biru kemudian -> isi `KETERANGAN = "Data cocok"` pada file sertifikat.

### **Special Case Pada Keterangan**
-Jika terdapat dua keterangan atau lebih pada satu NIB, maka tolong pada kolom keterangan diberi penomoran. 1,2,3 dan seterusnya. biar saya bisa tau telah terjadi update apa saja pada data tersebut


## Fix Error
<!-- 1. masih ada data double cek sesuaikan dengan file excell hasi; tidak ada NIB yang sama dalah satu file wajib pajak  -->
<!-- 2. Kalau update data cell abu-abu, hilangkan keterangan lama di sertifikat dan wajib pajak -->
<!-- 3. Jika cell abu-abu tidak cocok atau NIB tidak ada, ganti cell menjadi merah. jadi hasil akhir tidak da lagi cell berwarna abu-abu karena semua cell sudah di analisa. -->
<!-- 4. pengisian data hijau, sesuai logic exact_match -->
