# USER MANUAL - TAX ADMINISTRATION SYSTEM

## 📋 DAFTAR ISI
1. [Pengenalan Sistem](#pengenalan-sistem)
2. [Login](#login)
3. [Dashboard](#dashboard)
4. [Master Data](#master-data)
5. [Input Pajak](#input-pajak)
6. [Perhitungan](#perhitungan)
7. [Pelaporan](#pelaporan)
8. [SPT](#spt)
9. [Troubleshooting](#troubleshooting)

---

## Pengenalan Sistem

### Apa itu Tax Administration System?

Tax Administration System adalah aplikasi Excel berbasis VBA Macros yang dirancang untuk mengelola berbagai jenis pajak:

- **PPh OP** (Pajak Penghasilan Orang Pribadi)
- **PPh Badan** (Pajak Penghasilan Badan Usaha)
- **PPN** (Pajak Pertambahan Nilai)
- **PPh Pasal 21** (Pajak Penghasilan Karyawan)
- **PPh Pasal 23** (Dividen, Bunga, Royalti)
- **PPh Pasal 4 Ayat 2** (Pajak Final)
- **PPh Pasal 24** (Kredit Pajak Luar Negeri)
- **PPh Pasal 26** (Penghasilan Bukan Residans)
- **Unifikasi** (Rekapitulasi Semua Pajak)

### Keunggulan Sistem

✅ **User-Friendly**: Interface yang mudah digunakan, semua berbasis klik tombol  
✅ **Otomatis**: Perhitungan pajak dilakukan otomatis  
✅ **Terstruktur**: Data disimpan terorganisir dalam database internal Excel  
✅ **Secure**: Login system untuk kontrol akses  
✅ **Reporting**: Generate laporan & SPT otomatis  
✅ **Export**: Bisa export ke PDF & print  

---

## LOGIN

### Cara Login

1. **Buka file Excel** `Tax-Administration-Template.xlsm`
2. Akan muncul sheet **"LOGIN"** secara otomatis
3. Klik tombol berwarna **"LOGIN"**
4. Dialog box akan muncul meminta **Username**
5. Input username (default: `admin`)
6. Tekan Enter atau klik OK
7. Dialog kedua meminta **Password** (default: `123456`)
8. Input password
9. Tekan Enter atau klik OK

### Hasil Login

#### ✅ Jika Berhasil:
- Muncul pesan: "Selamat datang, [Username]!"
- Otomatis pindah ke sheet **"DASHBOARD"**
- Anda siap menggunakan sistem

#### ❌ Jika Gagal:
- Muncul pesan: "Username atau Password salah!"
- Option untuk login ulang
- Ulangi langkah 1-9

### Mengubah Username & Password

1. Minta akses ke file master (hubungi administrator)
2. Buka sheet "ADMIN" (hidden sheet)
3. Edit tabel user di kolom Username & Password
4. Perbarui macro `LoginUser()` jika perlu
5. Save file dengan format .xlsm

---

## DASHBOARD

### Overview Dashboard

Dashboard adalah menu utama yang menampilkan:

```
╔═════════════════════════════════════════════╗
║         TAX ADMINISTRATION SYSTEM            ║
║              MAIN DASHBOARD                  ║
╠═════════════════════════════════════════════╣
║                                             ║
║  📊 STATISTIK OVERVIEW                     ║
║  • Total Wajib Pajak: X                    ║
║  • Total Pajak Terutang: Rp X,XXX,XXX      ║
║                                             ║
║  🔘 MENU NAVIGASI:                         ║
║  [MASTER DATA] [INPUT PAJAK] [PERHITUNGAN] ║
║  [PELAPORAN] [SPT] [LOGOUT]                ║
║                                             ║
╚═════════════════════════════════════════════╝
```

### Tombol Navigasi

| Tombol | Fungsi | Keterangan |
|--------|--------|-----------|
| **MASTER DATA** | Manage Wajib Pajak & Tarif | Tambah, Edit, Hapus, Cari WP |
| **INPUT PAJAK** | Input data pajak | Pilih jenis pajak untuk input |
| **PERHITUNGAN** | Lihat rekapitulasi | Otomatis hitung semua pajak |
| **PELAPORAN** | Generate laporan | Export & print laporan |
| **SPT** | Buat SPT | Generate SPT format DJP |
| **LOGOUT** | Keluar sistem | Kembali ke login |

---

## MASTER DATA

### Submenu Master Data

Klik tombol **MASTER DATA** di Dashboard, akan muncul menu:

```
Pilih opsi:
1 = Lihat Daftar WP
2 = Tambah WP
3 = Edit WP
4 = Hapus WP
5 = Cari WP
```

### 1️⃣ LIHAT DAFTAR WP (Opsi 1)

**Fungsi**: Menampilkan semua daftar Wajib Pajak

**Langkah**:
1. Pilih opsi **1** di menu Master Data
2. Sheet "DATA-WP" akan terbuka
3. Lihat tabel lengkap berisi:
   - NPWP
   - Nama Wajib Pajak
   - Jenis (OP/Badan)
   - Alamat
   - Status (Aktif/Tidak Aktif)
   - Tanggal Input

**Contoh Data**:
```
NPWP              | Nama           | Jenis | Status
05.123.456.7-890  | PT ABC Maju    | Badan | Aktif
05.234.567.8-901  | Tuan Budi      | OP    | Aktif
05.345.678.9-012  | CV Maju Jaya   | Badan | Aktif
```

### 2️⃣ TAMBAH WP (Opsi 2)

**Fungsi**: Menambah Wajib Pajak baru

**Langkah**:
1. Pilih opsi **2** di menu Master Data
2. Dialog pertama: "Masukkan NPWP (format: XX.XXX.XXX.X-XXX.XXX)"
   - Contoh: `05.123.456.7-890`
3. Dialog kedua: "Masukkan Nama Wajib Pajak"
   - Contoh: `PT ABC Indonesia`
4. Dialog ketiga: "Masukkan Jenis (OP/Badan)"
   - Ketik: `OP` atau `Badan`
5. Dialog keempat: "Masukkan Alamat"
   - Contoh: `Jl. Merdeka No. 123, Jakarta`
6. Dialog kelima: "Masukkan Status (Aktif/Tidak Aktif)"
   - Ketik: `Aktif` atau `Tidak Aktif`
7. Klik OK
8. Pesan sukses akan muncul: "Data Wajib Pajak berhasil ditambahkan!"

**Catatan**: Data otomatis disimpan dengan timestamp (tanggal & jam input)

### 3️⃣ EDIT WP (Opsi 3)

**Fungsi**: Mengubah data Wajib Pajak yang sudah ada

**Langkah**:
1. Pilih opsi **3** di menu Master Data
2. Dialog: "Masukkan NPWP yang ingin diedit"
   - Contoh: `05.123.456.7-890`
3. Jika ditemukan, akan muncul dialog untuk setiap field:
   - "Edit Nama (tekan Enter jika tidak diubah)"
   - "Edit Jenis OP/Badan (tekan Enter jika tidak diubah)"
   - "Edit Alamat (tekan Enter jika tidak diubah)"
   - "Edit Status (tekan Enter jika tidak diubah)"
4. Isi field yang ingin diubah, biarkan kosong untuk field yang tidak diubah
5. Klik OK
6. Pesan sukses: "Data Wajib Pajak berhasil diubah!"

**Contoh**:
- NPWP ditemukan: `05.123.456.7-890`
- Ingin ubah nama dari `PT ABC Lama` → `PT ABC Baru`
- Ketik: `PT ABC Baru` di field Nama, tekan Enter untuk field lain
- Data tersimpan

### 4️⃣ HAPUS WP (Opsi 4)

**Fungsi**: Menghapus data Wajib Pajak

**Langkah**:
1. Pilih opsi **4** di menu Master Data
2. Dialog: "Masukkan NPWP yang ingin dihapus"
3. Konfirmasi: "Yakin ingin menghapus data ini?"
   - Klik **Yes** untuk hapus
   - Klik **No** untuk batal
4. Pesan: "Data berhasil dihapus!" atau "Penghapusan dibatalkan"

### 5️⃣ CARI WP (Opsi 5)

**Fungsi**: Mencari Wajib Pajak berdasarkan NPWP atau Nama

**Langkah**:
1. Pilih opsi **5** di menu Master Data
2. Dialog: "Masukkan NPWP atau Nama yang dicari"
   - Bisa search by NPWP: `05.123.456.7-890`
   - Atau search by Nama: `ABC` (akan search partial match)
3. Jika ditemukan, akan menampilkan:
   ```
   Ditemukan:
   NPWP: 05.123.456.7-890
   Nama: PT ABC Indonesia
   Jenis: Badan
   Alamat: Jl. Merdeka No. 123, Jakarta
   Status: Aktif
   ```
4. Row yang ditemukan akan ter-highlight

---

## INPUT PAJAK

### Menu Input Pajak

Klik tombol **INPUT PAJAK** di Dashboard, akan muncul menu:

```
Pilih jenis pajak:
1 = PPh OP
2 = PPN
3 = PPh Pasal 21
4 = PPh Pasal 23
5 = Kembali
```

### 1️⃣ INPUT PPh OP (Pajak Penghasilan Orang Pribadi)

**Deskripsi**: Pajak dari penghasilan individu/orang pribadi

**Langkah Input**:
1. Pilih opsi **1** - INPUT PPh OP
2. Dialog 1: "Masukkan NPWP"
   - Contoh: `05.123.456.7-890`
3. Dialog 2: "Masukkan Nama WP"
   - Contoh: `Tuan Budi Santoso`
4. Dialog 3: "Masukkan Penghasilan Bruto"
   - Contoh: `500000000` (Rp 500 juta)
5. Dialog 4: "Masukkan Biaya (pengurang)"
   - Contoh: `100000000` (Rp 100 juta)
6. Sistem otomatis menghitung PPh berdasarkan bracket tarif

**Perhitungan Otomatis**:
```
Penghasilan Bruto: Rp 500.000.000
Biaya Pengurang:   Rp 100.000.000
─────────────────────────────────
PKP (Kena Pajak):  Rp 400.000.000

Tarif Bracket (2024):
- 0 - 60M    : 5%
- 60M - 250M : 15%
- 250M - 500M: 25%
- >500M      : 30%

Perhitungan:
- Bracket 1: 60M × 5%     = Rp 3.000.000
- Bracket 2: 190M × 15%   = Rp 28.500.000
- Bracket 3: 150M × 25%   = Rp 37.500.000
─────────────────────────────────
PPh OP Terutang: Rp 69.000.000
```

**Hasil**: Data tersimpan dengan PPh yang telah dihitung, muncul pesan "PPh OP: Rp 69.000.000"

### 2️⃣ INPUT PPN (Pajak Pertambahan Nilai)

**Deskripsi**: Pajak atas transaksi penjualan barang/jasa

**Langkah Input**:
1. Pilih opsi **2** - INPUT PPN
2. Dialog 1: "Masukkan NPWP Penjual"
3. Dialog 2: "Masukkan DPP (Dasar Pengenaan Pajak)"
   - Contoh: `100000000` (Rp 100 juta)
4. Dialog 3: "Masukkan Rate PPN % (default 11)"
   - Default: `11` (bisa diganti dengan rate yang berlaku)

**Perhitungan Otomatis**:
```
DPP:       Rp 100.000.000
Rate PPN:  11%
─────────────────────────
PPN:       Rp 11.000.000
```

### 3️⃣ INPUT PPh PASAL 21 (Pajak Penghasilan Karyawan)

**Deskripsi**: Pajak atas gaji/upah karyawan

**Langkah Input**:
1. Pilih opsi **3** - INPUT PPh Pasal 21
2. Dialog 1: "Masukkan NPWP Karyawan"
3. Dialog 2: "Masukkan Nama Karyawan"
4. Dialog 3: "Masukkan Gaji Pokok"
   - Contoh: `10000000` (Rp 10 juta)
5. Dialog 4: "Masukkan Tunjangan"
   - Contoh: `2000000` (Rp 2 juta)
   - Tekan 0 jika tidak ada
6. Dialog 5: "Masukkan Potongan (JKK, JKM, dll)"
   - Contoh: `500000` (Rp 500 ribu)
   - Tekan 0 jika tidak ada

**Perhitungan Otomatis**:
```
Gaji Pokok:        Rp 10.000.000
Tunjangan:         Rp  2.000.000
Potongan:          Rp    500.000
─────────────────────────────────
Penghasilan Neto:  Rp 11.500.000

PTKP (2024):       Rp 54.000.000
─────────────────────────────────
PKP:               Rp  0 (Kurang dari PTKP)

PPh Pasal 21:      Rp  0 (Tidak ada pajak)
```

### 4️⃣ INPUT PPh PASAL 23

**Status**: Fitur sedang dikembangkan

---

## PERHITUNGAN

### Fungsi Perhitungan

Menu **PERHITUNGAN** menampilkan rekapitulasi semua pajak yang telah diinput.

**Langkah**:
1. Klik tombol **PERHITUNGAN** di Dashboard
2. Sistem otomatis generate laporan
3. Sheet baru **"REPORT-SUMMARY"** akan terbuat/terbuka

**Tampilan Laporan**:
```
LAPORAN PAJAK TERUTANG
Tanggal: 25-05-2026

Jenis Pajak                 Jumlah (Rp)
─────────────────────────────────────
PPh OP                      69.000.000
PPN                         11.000.000
PPh Pasal 21                     0
─────────────────────────────────────
TOTAL PAJAK TERUTANG       80.000.000
```

**Fitur**:
- ✅ Rekapitulasi otomatis
- ✅ Bisa print
- ✅ Bisa export ke PDF
- ✅ Update realtime setiap kali ada input baru

---

## PELAPORAN

### Fungsi Pelaporan

Menu **PELAPORAN** untuk generate dan export laporan pajak.

**Langkah**:
1. Klik tombol **PELAPORAN** di Dashboard
2. Sistem generate laporan otomatis
3. Sheet "REPORT-SUMMARY" akan terbuka
4. Pilih aksi:
   - 📄 **PRINT**: Ctrl+P atau Ctrl+Shift+K
   - 📤 **EXPORT PDF**: Klik tombol "Export ke PDF"

### Tombol di Laporan

#### Print Laporan

**Langkah**:
1. Di sheet laporan, tekan **Ctrl+P**
2. Atau klik tombol "PRINT"
3. Print preview akan muncul
4. Atur setting printer
5. Klik **Print**

#### Export ke PDF

**Langkah**:
1. Di sheet laporan, klik tombol **"EXPORT ke PDF"**
2. Dialog: "Masukkan nama file (tanpa .pdf)"
   - Contoh: `Laporan_Pajak_Mei_2026`
3. Klik OK
4. File PDF otomatis tersimpan di Desktop dengan nama: `Laporan_Pajak_Mei_2026.pdf`
5. Muncul pesan sukses

---

## SPT

### Fungsi SPT

Menu **SPT** untuk membuat Surat Pemberitahuan sesuai format DJP (Direktorat Jenderal Pajak).

**Langkah**:
1. Klik tombol **SPT** di Dashboard
2. Dialog: "Masukkan NPWP untuk membuat SPT"
   - Contoh: `05.123.456.7-890`
3. Sistem generate SPT otomatis
4. Sheet baru **"SPT-[NPWP]"** akan terbuat
   - Contoh: `SPT-05.123.456.7-890`

### Struktur SPT

```
SURAT PEMBERITAHUAN (SPT)

NPWP: 05.123.456.7-890
Tanggal Laporan: 25-05-2026
Periode: May 2026

PENGHASILAN
Penghasilan dari Usaha/Bisnis:    Rp [isi]
Penghasilan dari Pekerjaan:       Rp [isi]
Penghasilan dari Modal:           Rp [isi]
Penghasilan Lainnya:              Rp [isi]
─────────────────────────────────────────
Jumlah Penghasilan Bruto:         Rp [isi]

PENGURANG
Biaya Operasional:                Rp [isi]
Biaya Lainnya:                    Rp [isi]
─────────────────────────────────────────
Jumlah Pengurang:                 Rp [isi]

PENGHASILAN KENA PAJAK:           Rp [isi]

PERHITUNGAN PAJAK
PPh Terutang:                     Rp [isi]
PPh Yang Sudah Dibayar:           Rp [isi]
Kurang / Lebih Bayar:             Rp [isi]
```

### Cara Isi SPT

1. **Section PENGHASILAN**: Isi semua jenis penghasilan
2. **Section PENGURANG**: Isi semua jenis pengurang/biaya
3. **Sistem otomatis**: Hitung PKP dan PPh terutang
4. **Section PERHITUNGAN PAJAK**: Lihat PPh terutang vs yang sudah dibayar

---

## LOGOUT

### Cara Logout

1. Klik tombol **LOGOUT** di Dashboard
2. Muncul pesan: "Anda telah logout. Silakan login kembali."
3. Otomatis kembali ke sheet **"LOGIN"**
4. Untuk login kembali, ulangi langkah di section [LOGIN](#login)

---

## TROUBLESHOOTING

### ❌ Masalah: Macro tidak berjalan saat klik tombol

**Solusi**:
1. Pastikan sudah klik "Enable Macros" saat membuka file
2. Jika tidak ada notifikasi, lakukan:
   - Klik **File** → **Info**
   - Jika ada peringatan "Macros have been disabled", klik **Enable Macros**
3. Jika masih tidak jalan, setup manual:
   - Klik **Developer** → **Visual Basic** (Ctrl+Alt+F11)
   - Cek apakah ada kode di Module1
   - Jika tidak ada, copy-paste dari `Macro-Code-Complete.vba`

### ❌ Masalah: Error "Sub or Function not defined"

**Solusi**:
1. Buka Developer → Visual Basic
2. Lihat di panel kiri, apakah ada Module dengan kode
3. Jika tidak ada, Insert Module baru
4. Copy-paste SELURUH kode dari `Macro-Code-Complete.vba`
5. Pastikan tidak ada bagian yang terpotong
6. Save file dengan format .xlsm

### ❌ Masalah: Tidak bisa login

**Solusi**:
1. Cek username: `admin` (perhatikan huruf besar/kecil)
2. Cek password: `123456`
3. Pastikan keyboard dalam bahasa Inggris
4. Jika lupa password, hubungi administrator
5. Alternatif: Edit di sheet "ADMIN" (jika punya akses)

### ❌ Masalah: Dropdown NPWP kosong saat input pajak

**Solusi**:
1. Pastikan sudah tambah data Wajib Pajak terlebih dahulu
2. Cek sheet "DATA-WP" sudah ada data
3. Refresh data: Menu → Data → Refresh All
4. Tutup file & buka lagi
5. Jika masih tidak ada, setup dropdown di form input

### ❌ Masalah: Excel freeze/tidak responsive

**Solusi**:
1. Tekan **Ctrl + Break** untuk stop macro yang sedang jalan
2. Tunggu beberapa detik
3. Jika masih freeze, restart Excel
4. Buka kembali file
5. Jika berulang, cek logic di macro (mungkin ada loop tak terbatas)

### ❌ Masalah: Data tidak tersimpan

**Solusi**:
1. Cek apakah sheet "DATA-*" ada (tidak hidden/terhapus)
2. Klik menu **View** → **Unhide** untuk unhide sheet yang tersembunyi
3. Pastikan file dalam format `.xlsm` (bukan .xlsx)
4. Save manual dengan Ctrl+S
5. Jangan close file tanpa save

### ❌ Masalah: Tidak bisa export ke PDF

**Solusi**:
1. Pastikan Excel terinstall dengan fitur PDF export
2. Cek folder Desktop memiliki permission write
3. Tutup semua file PDF yang sedang terbuka
4. Coba lagi
5. Jika masih error, print ke PDF via printer virtual

### ✅ Masalah Lain?

Jika masalah tidak terselesaikan:
1. Screenshoot error message
2. Catat langkah yang Anda lakukan
3. Hubungi developer/administrator
4. Siapkan file Excel untuk debug

---

## QUICK REFERENCE

### Login Credentials
```
Username: admin
Password: 123456
```

### Shortcut Keyboard
```
Ctrl+Alt+F11     → Buka VBA Editor
Ctrl+P           → Print
Ctrl+S           → Save
Ctrl+Break       → Stop Macro
Ctrl+Shift+K     → Print Preview
```

### File Structure
```
Tax-Administration-Template.xlsm
├── LOGIN (Sheet login)
├── DASHBOARD (Menu utama)
├── DATA-WP (Master Wajib Pajak)
├── MASTER-TARIF (Master Tarif Pajak)
├── INPUT-PPh-OP (Form input PPh OP)
├── INPUT-PPN (Form input PPN)
├── INPUT-PPh21 (Form input PPh 21)
├── DATA-PPH-OP (Penyimpanan data PPh OP)
├── DATA-PPN (Penyimpanan data PPN)
├── DATA-PPH21 (Penyimpanan data PPh 21)
├── REPORT-SUMMARY (Laporan Pajak)
├── SPT-* (Sheet SPT per NPWP)
└── ADMIN (Hidden - Konfigurasi sistem)
```

---

## KONTAK SUPPORT

Untuk bantuan lebih lanjut:
- 📧 Email: [administrator email]
- 📱 WhatsApp: [administrator phone]
- 📞 Telepon: [administrator phone]
- ⏰ Jam Kerja: Senin-Jumat 08:00-17:00

---

**Semoga panduan ini membantu! Happy Taxing! 🎉**
