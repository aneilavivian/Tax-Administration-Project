# 📊 Tax Administration Project

**Sistem Administrasi Perpajakan Berbasis Excel VBA Macros**

> Prototype administrasi perpajakan untuk mengelola berbagai jenis pajak dengan interface yang user-friendly, perhitungan otomatis, dan reporting terintegrasi.

---

## 🎯 Project Overview

Sistem ini dirancang untuk memudahkan manajemen dan administrasi berbagai jenis pajak Indonesia melalui aplikasi Excel dengan fitur Macros. Semua pengguna hanya perlu melakukan klik-klik sederhana tanpa perlu memahami coding.

### Jenis Pajak yang Dikelola

✅ **PPh OP** - Pajak Penghasilan Orang Pribadi  
✅ **PPh Badan** - Pajak Penghasilan Badan Usaha  
✅ **PPN** - Pajak Pertambahan Nilai  
✅ **PPh Pasal 21** - Pajak Penghasilan Karyawan  
✅ **PPh Pasal 23** - Dividen, Bunga, Royalti  
✅ **PPh Pasal 4 Ayat 2** - Pajak Final  
✅ **PPh Pasal 24** - Kredit Pajak Luar Negeri  
✅ **PPh Pasal 26** - Penghasilan Bukan Residans  
✅ **Unifikasi** - Rekapitulasi Semua Pajak  

---

## ✨ Fitur Utama

### 🔐 **Sistem Keamanan**
- Login dengan username & password
- User access control
- Data administrator terlindungi

### 📋 **Master Data Management**
- Kelola data Wajib Pajak (OP/Badan)
- Manage tarif pajak
- CRUD operations (Tambah, Edit, Hapus, Cari)

### 💰 **Input & Perhitungan Otomatis**
- Form input pajak per jenis
- Perhitungan pajak otomatis sesuai tarif & ketentuan
- Validasi data real-time
- Penyimpanan data terstruktur

### 📊 **Reporting & Analytics**
- Generate laporan pajak terutang
- Rekapitulasi per jenis pajak
- Export ke PDF
- Print-ready format

### 📄 **SPT Generation**
- Buat Surat Pemberitahuan format DJP
- Auto-fill dari data input
- Export & print

### 🖥️ **User Interface**
- Dashboard interaktif
- Menu navigasi yang intuitif
- Semua berbasis klik tombol (point & click)
- Error handling & pesan informatif

---

## 📁 Struktur File

```
Tax-Administration-Project/
├── README.md                          # Dokumentasi ini
├── SETUP-INSTRUCTIONS.md              # Panduan setup macro
├── USER-MANUAL.md                     # Manual lengkap pengguna
├── Macro-Code-Complete.vba            # Kode VBA lengkap
├── Tax-Administration-Template.xlsm   # File Excel template (jika ada)
└── docs/
    ├── ARCHITECTURE.md                # Arsitektur sistem
    ├── INSTALLATION.md                # Instalasi detil
    └── TROUBLESHOOTING.md             # Panduan troubleshoot
```

---

## 🚀 Quick Start

### Prerequisites
- Microsoft Excel Desktop (2016 atau lebih baru)
- File permission untuk read/write

### Installation

#### Langkah 1: Download File
```bash
git clone https://github.com/aneilavivian/Tax-Administration-Project.git
cd Tax-Administration-Project
```

#### Langkah 2: Setup Macro
1. Buka file `Tax-Administration-Template.xlsm`
2. Buka **Developer** tab → **Visual Basic** (Ctrl+Alt+F11)
3. **Insert** → **Module**
4. Copy-paste seluruh kode dari `Macro-Code-Complete.vba`
5. **Save** file dengan format .xlsm

#### Langkah 3: Enable Macros
1. Close file Excel
2. Buka kembali `Tax-Administration-Template.xlsm`
3. Klik **Enable Macros** saat muncul peringatan

#### Langkah 4: Login
1. Sheet "LOGIN" akan terbuka otomatis
2. Klik tombol **LOGIN**
3. Username: `admin`
4. Password: `123456`
5. Tekan Enter

**Selesai! ✅ Sistem siap digunakan.**

---

## 📖 Dokumentasi Lengkap

### Setup & Installation
👉 **[SETUP-INSTRUCTIONS.md](SETUP-INSTRUCTIONS.md)**
- Panduan instalasi macro step-by-step
- Troubleshooting umum
- Konfigurasi awal

### User Manual
👉 **[USER-MANUAL.md](USER-MANUAL.md)**
- Panduan penggunaan lengkap
- Cara login
- Master data management
- Input pajak
- Reporting & SPT
- FAQ & Troubleshooting

### VBA Code
👉 **[Macro-Code-Complete.vba](Macro-Code-Complete.vba)**
- Kode VBA lengkap
- Penjelasan setiap function
- Ready to copy-paste

---

## 🎮 Cara Menggunakan

### 1. Login
```
Klik tombol LOGIN → Input username (admin) & password (123456)
```

### 2. Master Data
```
Dashboard → MASTER DATA
├─ Tambah Wajib Pajak
├─ Edit Wajib Pajak  
├─ Hapus Wajib Pajak
└─ Cari Wajib Pajak
```

### 3. Input Pajak
```
Dashboard → INPUT PAJAK
├─ PPh OP
├─ PPN
├─ PPh Pasal 21
└─ PPh Pasal 23
```

### 4. Perhitungan
```
Dashboard → PERHITUNGAN
→ Otomatis generate laporan rekapitulasi
```

### 5. Pelaporan
```
Dashboard → PELAPORAN
├─ Print laporan
└─ Export ke PDF
```

### 6. SPT
```
Dashboard → SPT
→ Input NPWP → Generate SPT format DJP
```

---

## 📊 Contoh Data Flow

```
┌─────────────────┐
│  LOGIN SYSTEM   │  Username: admin, Password: 123456
└────────┬────────┘
         │
         ▼
┌─────────────────────────┐
│   MASTER DATA INPUT     │  Wajib Pajak, Tarif Pajak
└────────┬────────────────┘
         │
         ▼
┌─────────────────────────┐
│   INPUT PAJAK DATA      │  PPh OP, PPN, PPh 21, dll
│  (Auto Calculation)     │
└────────┬────────────────┘
         │
         ▼
┌─────────────────────────┐
│   DATABASE STORAGE      │  Sheet "DATA-*" (Hidden)
└────────┬────────────────┘
         │
         ▼
┌─────────────────────────┐
│   PERHITUNGAN & REPORT  │  Summary pajak terutang
└────────┬────────────────┘
         │
         ├─▶ PRINT
         ├─▶ PDF EXPORT
         └─▶ SPT GENERATION
```

---

## 🔧 Fitur Macro Tersedia

### Authentication
| Macro | Fungsi |
|-------|--------|
| `LoginUser()` | Login user dengan validasi |
| `Logout()` | Logout & kembali ke login |

### Master Data
| Macro | Fungsi |
|-------|--------|
| `AddWajibPajak()` | Tambah wajib pajak baru |
| `EditWajibPajak()` | Edit data wajib pajak |
| `DeleteWajibPajak()` | Hapus data wajib pajak |
| `SearchWajibPajak()` | Cari wajib pajak |

### Calculation
| Macro | Fungsi |
|-------|--------|
| `CalculatePPhOP()` | Hitung PPh OP |
| `CalculatePPhBadan()` | Hitung PPh Badan |
| `CalculatePPN()` | Hitung PPN |
| `CalculatePPh21()` | Hitung PPh Pasal 21 |
| `CalculatePPh23()` | Hitung PPh Pasal 23 |

### Input Data
| Macro | Fungsi |
|-------|--------|
| `InputDataPPh_OP()` | Input PPh OP |
| `InputDataPPN()` | Input PPN |
| `InputDataPPh21()` | Input PPh Pasal 21 |

### Reporting
| Macro | Fungsi |
|-------|--------|
| `GenerateReportPajakTertanggung()` | Generate laporan |
| `GenerateSPT()` | Generate SPT |
| `PrintCurrentSheet()` | Print sheet |
| `ExportToPDF()` | Export ke PDF |

### Navigation
| Macro | Fungsi |
|-------|--------|
| `OpenDashboard()` | Buka dashboard |
| `OpenMasterData()` | Menu master data |
| `OpenInputMenu()` | Menu input pajak |
| `OpenPerhitungan()` | Buka perhitungan |
| `OpenSPT()` | Buka SPT |

---

## 🐛 Troubleshooting

### Macro tidak berjalan
**Solusi**: Enable macros di Excel (File → Info → Enable Macros)

### Error "Sub or Function not defined"
**Solusi**: Pastikan kode VBA sudah di-copy ke Module dengan benar

### Data tidak tersimpan
**Solusi**: Pastikan file dalam format .xlsm dan save dengan Ctrl+S

**Lebih lengkap:** Lihat [USER-MANUAL.md](USER-MANUAL.md#troubleshooting)

---

## 💡 Tips & Tricks

1. **Shortcut**: Tekan `Ctrl+Alt+F11` untuk membuka VBA Editor
2. **Debug**: Gunakan `Ctrl+Break` untuk stop macro yang hang
3. **Backup**: Selalu backup file sebelum modifikasi macro
4. **Format**: Gunakan format .xlsm (macro-enabled workbook)
5. **Testing**: Test dengan data sample sebelum data production

---

## 📝 Peraturan & Tarif

Sistem menggunakan tarif pajak **2024** sesuai ketentuan DJP Indonesia:

- **PPh OP**: Bracket 5%, 15%, 25%, 30%
- **PPh Badan**: 20%
- **PPN**: 11% (dapat disesuaikan)
- **PPh Pasal 21**: Progresif sesuai PKP
- **PPh Pasal 23**: 10% (Dividen), 15% (Bunga), 15% (Royalti), 2% (Jasa)

**Catatan**: Pastikan untuk update tarif sesuai peraturan terbaru di sheet "MASTER-TARIF"

---

## 🤝 Kontribusi

Untuk berkontribusi ke project ini:

1. Fork repository
2. Buat branch feature (`git checkout -b feature/NamaFeature`)
3. Commit changes (`git commit -m 'Tambah fitur xxx'`)
4. Push ke branch (`git push origin feature/NamaFeature`)
5. Buat Pull Request

---

## 📋 Roadmap

### ✅ Done
- [x] Setup struktur sistem
- [x] Login & authentication
- [x] Master data management
- [x] Input pajak (PPh OP, PPN, PPh 21)
- [x] Auto calculation
- [x] Reporting & export PDF
- [x] SPT generation
- [x] User manual lengkap

### 🔄 In Progress
- [ ] Dashboard dengan chart/grafik
- [ ] Input PPh Badan, Pasal 23, dll
- [ ] Email integration untuk notifikasi
- [ ] Export ke Excel dengan format lebih baik
- [ ] Web version (future)

### 📋 Todo
- [ ] Multi-user environment
- [ ] Database integration (SQL)
- [ ] Mobile app version
- [ ] API development
- [ ] Cloud sync

---

## 📞 Support & Contact

| Kontak | Info |
|--------|------|
| 📧 Email | aneilavivian@gmail.com |
| 📁 GitHub | [aneilavivian](https://github.com/aneilavivian) |
| 📱 Repository | [Tax-Administration-Project](https://github.com/aneilavivian/Tax-Administration-Project) |

---

## 📜 License

Project ini dibuat untuk keperluan tugas mata kuliah. Silakan gunakan & modifikasi sesuai kebutuhan.

---

## 📚 Referensi

- [Direktorat Jenderal Pajak (DJP)](https://www.pajak.go.id)
- [Peraturan Pajak Indonesia](https://www.pajak.go.id/en/regulation)
- [Microsoft Excel VBA Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview)
- [Excel Macro Tutorial](https://www.excel-easy.com/vba.html)

---

## 🙏 Terima Kasih

Terima kasih telah menggunakan **Tax Administration Project**!

Semoga sistem ini membantu dalam mengelola administrasi perpajakan dengan lebih efisien. 🎉

---

**Last Updated**: 25 Mei 2026  
**Version**: 1.0  
**Status**: ✅ Ready for Use
