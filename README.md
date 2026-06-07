# 📊 Kumpulan Algoritma Data Mining Menggunakan Microsoft Excel

> Implementasi algoritma Data Mining — K-Means Clustering, Naïve Bayes Classification, dan K-Nearest Neighbor (KNN) — langsung di atas spreadsheet Microsoft Excel menggunakan formula native dan VBA Macro.

---

## 📋 Daftar Isi

- [Deskripsi Proyek](#deskripsi-proyek)
- [Daftar Algoritma](#daftar-algoritma)
- [Mekanisme Deployment](#mekanisme-deployment-teknis)
- [Cara Menggunakan](#cara-menggunakan)
- [Persyaratan Sistem](#persyaratan-sistem)

---

## 📖 Deskripsi Proyek

Repositori ini berisi kumpulan implementasi algoritma **Data Mining** yang dibangun sepenuhnya di atas **Microsoft Excel** — tanpa bahasa pemrograman eksternal, tanpa framework, tanpa server. Setiap file Excel berfungsi sekaligus sebagai **antarmuka**, **engine kalkulasi**, dan **media visualisasi** hasil komputasi algoritma.

Pendekatan ini membuktikan bahwa spreadsheet bukan sekadar alat tabel biasa, melainkan dapat menjadi platform komputasi ilmiah yang powerful dengan memanfaatkan formula bawaan Excel, array formula, dan otomasi melalui VBA Macro.

---

## 🗂️ Daftar Algoritma

| No | File | Format | Algoritma | Teknik Excel yang Digunakan |
|---|---|---|---|---|
| 1 | `K-means + Latihan.xlsx` | Excel Workbook | **K-Means Clustering** | Formula native + Array Formula |
| 2 | `Naives Bayes.xlsm` | Excel Macro-Enabled | **Naïve Bayes Classification** | Formula + **VBA Macro** |
| 3 | `Tabel 1-5 KNN.xlsx` | Excel Workbook | **K-Nearest Neighbor (KNN)** | Formula native + Array Formula |

---

## 🚀 Mekanisme Deployment (Teknis)

Ini adalah bagian terpenting untuk dipahami: proyek ini menggunakan paradigma **"Spreadsheet-as-Computation"** — sebuah pendekatan di mana file Excel itu sendiri **adalah** aplikasinya, bukan sekadar media penyimpan data.

---

### 1. Tidak Ada Build, Tidak Ada Server

Berbeda dari aplikasi konvensional seperti Android (perlu di-compile dengan Gradle) atau aplikasi desktop (perlu di-build dengan MSBuild), proyek ini **tidak memiliki proses build sama sekali**.

```
Aplikasi Konvensional:
  Source Code → [Compile/Build] → Executable → Deploy ke Server/Device

Proyek Ini (Spreadsheet-as-Computation):
  File .xlsx / .xlsm → [Buka di Excel] → Langsung berjalan
```

File yang ada di repositori ini **langsung siap pakai** begitu dibuka di Microsoft Excel. Tidak ada dependensi yang perlu diinstal, tidak ada environment yang perlu dikonfigurasi, tidak ada runtime yang perlu disiapkan — selain Microsoft Excel itu sendiri.

---

### 2. Excel sebagai Engine Komputasi

Ketika pengguna membuka file dan mengisi data, Microsoft Excel berperan sebagai **mesin eksekusi** yang menjalankan seluruh logika algoritma secara real-time.

#### a. Formula Native sebagai Kode Program

Pada file `.xlsx` (K-Means dan KNN), seluruh algoritma diimplementasikan menggunakan **formula bawaan Excel**. Formula-formula ini berperan layaknya fungsi dalam bahasa pemrograman:

```
Contoh formula yang digunakan dalam implementasi algoritma:

SQRT()          → Menghitung jarak Euclidean antar data point
SUMIF()         → Menjumlahkan nilai berdasarkan kondisi cluster
COUNTIF()       → Menghitung anggota dalam satu cluster
AVERAGE()       → Menghitung centroid baru tiap iterasi K-Means
MIN()           → Menentukan jarak terdekat pada KNN
INDEX(MATCH())  → Lookup label kelas pada KNN
Array Formula   → Komputasi vektor/matriks secara serentak (Ctrl+Shift+Enter)
```

Setiap kali nilai input berubah, Excel secara otomatis **me-recalculate** seluruh rantai formula yang bergantung padanya — ini yang disebut **Dependency Chain Recalculation**, mekanisme inti yang membuat spreadsheet dapat "menjalankan" algoritma secara live tanpa perlu tombol "Run".

#### b. VBA Macro sebagai Otomasi Proses

Pada file `Naives Bayes.xlsm`, selain formula Excel, terdapat **VBA Macro (Visual Basic for Applications)** yang tersimpan di dalam file itu sendiri. VBA adalah bahasa pemrograman embedded di Microsoft Excel yang memungkinkan otomasi proses yang terlalu kompleks untuk ditulis sebagai formula biasa.

```
Naives Bayes.xlsm
    │
    ├── Sheet Data       → Tempat input dataset training
    ├── Sheet Perhitungan → Tabel probabilitas prior & likelihood
    ├── Sheet Prediksi   → Input data uji & output klasifikasi
    └── VBA Module (tersembunyi di dalam file)
          │
          └── Sub/Function VBA yang mengotomasi:
                ├── Kalkulasi probabilitas P(C) untuk setiap kelas
                ├── Kalkulasi likelihood P(Xi|C) per fitur per kelas
                ├── Perkalian probabilitas (Naive Bayes theorem)
                └── Penentuan kelas prediksi berdasarkan P(C|X) tertinggi
```

VBA Macro dipicu melalui tombol yang tersedia di sheet, bukan berjalan otomatis — ini adalah **event-driven execution** dalam konteks spreadsheet.

---

### 3. Anatomi File Excel sebagai Aplikasi

Setiap file Excel dalam repositori ini memiliki struktur multi-sheet yang terorganisasi layaknya arsitektur aplikasi berlapis:

```
File Excel (satu file = satu aplikasi algoritma)
│
├── Sheet "Data / Dataset"
│     └── Tempat input data mentah (fitur, label, atribut)
│           ↕ (direferensikan oleh sheet berikutnya)
│
├── Sheet "Proses / Perhitungan"
│     └── Seluruh tahap komputasi algoritma berjalan di sini
│         menggunakan formula Excel yang saling berantai
│           ↕ (mengambil data dari sheet atas, meneruskan ke sheet bawah)
│
├── Sheet "Hasil / Output"
│     └── Menampilkan hasil akhir: cluster, label kelas, ranking jarak
│
└── Sheet "Latihan / Studi Kasus" (khusus K-Means)
      └── Contoh soal latihan dengan dataset berbeda
```

---

### 4. Format File: `.xlsx` vs `.xlsm`

Pemilihan format file bukan hal sembarangan — masing-masing punya perbedaan teknis penting:

| Aspek | `.xlsx` (K-Means, KNN) | `.xlsm` (Naïve Bayes) |
|---|---|---|
| Kepanjangan | Excel Open XML Workbook | Excel Macro-Enabled Workbook |
| Standar | ISO/IEC 29500 (OOXML) | ISO/IEC 29500 + VBA extension |
| Berisi VBA Macro | ❌ Tidak | ✅ Ya |
| Ukuran file | Lebih kecil | Lebih besar (menyimpan kode VBA) |
| Keamanan | Aman dibuka langsung | Muncul peringatan macro saat dibuka |
| Cocok untuk | Algoritma berbasis formula murni | Algoritma yang butuh otomasi/looping |

> ⚠️ Saat membuka file `.xlsm`, Excel akan menampilkan **security warning** tentang macro. Klik **"Enable Content"** agar VBA Macro dapat berjalan dan algoritma bekerja dengan benar.

---

### 5. Distribusi & Portabilitas

Salah satu keunggulan utama pendekatan ini dari sisi deployment adalah **portabilitasnya yang luar biasa**:

```
Cara "deploy" proyek ini:

GitHub Repository
      │
      │  (download / git clone)
      ▼
File .xlsx / .xlsm di komputer pengguna
      │
      │  (dobel klik)
      ▼
Microsoft Excel membuka file
      │
      │  (semua formula & macro sudah ada di dalam file)
      ▼
Algoritma langsung bisa digunakan
```

Tidak diperlukan:
- ❌ Instalasi runtime atau SDK
- ❌ Konfigurasi environment
- ❌ Koneksi database atau internet
- ❌ Proses compile atau build
- ❌ Dependensi library eksternal

Yang diperlukan hanya satu: **Microsoft Excel** (versi 2016 ke atas direkomendasikan).

---

### 6. Alur Eksekusi Algoritma di Dalam Excel

Berikut adalah bagaimana masing-masing algoritma "berjalan" secara teknis di dalam Excel:

#### K-Means Clustering (K-means + Latihan.xlsx)
```
Input: Dataset (n baris × m fitur)
  │
  ▼
Inisialisasi centroid awal (dipilih manual atau random)
  │
  ▼
Hitung jarak Euclidean tiap data ke tiap centroid
  [Formula: =SQRT(SUMXMY2(data_range, centroid_range))]
  │
  ▼
Assign tiap data ke cluster dengan centroid terdekat
  [Formula: =INDEX(label_cluster, MATCH(MIN(jarak), jarak, 0))]
  │
  ▼
Hitung centroid baru = rata-rata anggota tiap cluster
  [Formula: =AVERAGEIF(cluster_col, label, fitur_col)]
  │
  ▼
Ulangi secara manual (update centroid → recalculate otomatis)
  │
  ▼
Output: Label cluster tiap data point + visualisasi
```

#### Naïve Bayes Classification (Naives Bayes.xlsm)
```
Input: Dataset training (fitur + label kelas)
  │
  ▼
[VBA Macro] Hitung probabilitas prior P(C)
  P(C) = jumlah_data_kelas_C / total_data
  │
  ▼
[VBA Macro] Hitung likelihood P(Xi|C) untuk setiap fitur
  P(Xi=v|C) = jumlah_data_dengan_Xi=v_dan_kelas_C / jumlah_data_kelas_C
  │
  ▼
Input: Data baru yang ingin diklasifikasi
  │
  ▼
[Formula] Hitung P(C|X) = P(C) × ΠP(Xi|C)  untuk setiap kelas
  │
  ▼
Output: Kelas dengan nilai P(C|X) tertinggi = hasil prediksi
```

#### K-Nearest Neighbor (Tabel 1-5 KNN.xlsx)
```
Input: Dataset training + 1 data uji baru
  │
  ▼
Hitung jarak Euclidean antara data uji ke SEMUA data training
  [Formula: =SQRT((x1-a1)^2 + (x2-a2)^2 + ... + (xn-an)^2)]
  │
  ▼
Urutkan jarak dari terkecil ke terbesar
  [Formula: =SMALL(jarak_range, BARIS())]
  │
  ▼
Ambil K tetangga terdekat (nilai K ditentukan pengguna)
  │
  ▼
Hitung mayoritas label kelas dari K tetangga
  [Formula: =COUNTIF(label_K_tetangga, kelas_X)]
  │
  ▼
Output: Label kelas mayoritas = hasil klasifikasi data uji
```

---

## 🔧 Cara Menggunakan

### Langkah Umum

**1. Clone atau download repositori**
```bash
git clone https://github.com/Fahrozialdinata03/Kumpulan-Algoritma-Menggunakan-excel.git
```
Atau klik tombol **Code → Download ZIP** di halaman GitHub.

**2. Buka file yang diinginkan**
```
Dobel klik file .xlsx atau .xlsm
→ Microsoft Excel akan membuka file secara otomatis
```

**3. Khusus file `.xlsm` (Naïve Bayes)**
```
Saat muncul security bar di bagian atas Excel:
→ Klik "Enable Content" atau "Enable Macros"
→ Wajib dilakukan agar VBA Macro dapat berjalan
```

**4. Ikuti petunjuk di setiap sheet**
```
Setiap file memiliki sheet panduan/petunjuk penggunaan
Isi data pada kolom/sel yang telah disediakan (biasanya berwarna atau berlabel)
Hasil perhitungan akan muncul otomatis
```

---

### Per Algoritma

| Algoritma | File | Yang Perlu Dilakukan |
|---|---|---|
| K-Means | `K-means + Latihan.xlsx` | Isi dataset → set jumlah cluster (K) → update centroid per iterasi |
| Naïve Bayes | `Naives Bayes.xlsm` | Isi data training → klik tombol macro → input data uji → lihat hasil klasifikasi |
| KNN | `Tabel 1-5 KNN.xlsx` | Isi dataset training → input data uji → set nilai K → lihat ranking jarak & hasil |

---

## 💻 Persyaratan Sistem

| Komponen | Spesifikasi |
|---|---|
| Aplikasi | **Microsoft Excel 2016** atau lebih baru |
| Sistem Operasi | Windows 7/8/10/11 atau macOS (dengan Excel for Mac) |
| RAM | Minimum 2 GB (4 GB untuk dataset besar) |
| Macro Setting | **Macro harus diaktifkan** untuk file `.xlsm` |
| Alternatif | WPS Office atau LibreOffice Calc *(kompatibilitas tidak dijamin penuh)* |

> 💡 **Catatan:** Google Sheets **tidak direkomendasikan** karena beberapa formula array dan VBA Macro tidak kompatibel dengan platform berbasis web tersebut.

---

## 📚 Algoritma yang Diimplementasikan

### 1. K-Means Clustering
Algoritma unsupervised learning yang mengelompokkan data ke dalam K cluster berdasarkan kedekatan jarak ke centroid. Iterasi dilakukan hingga posisi centroid tidak berubah (konvergen).

### 2. Naïve Bayes Classification
Algoritma klasifikasi probabilistik berbasis Teorema Bayes dengan asumsi independensi antar fitur. Sangat efektif untuk klasifikasi teks dan data kategorikal.

### 3. K-Nearest Neighbor (KNN)
Algoritma klasifikasi berbasis instance yang mengklasifikasi data baru berdasarkan label mayoritas dari K tetangga terdekatnya di ruang fitur.

---

## 👨‍💻 Developer

**Fahrozi Aldinata**
- GitHub: [@Fahrozialdinata03](https://github.com/Fahrozialdinata03)

---

<p align="center">
  Dibuat dengan ❤️ menggunakan Microsoft Excel
</p>
