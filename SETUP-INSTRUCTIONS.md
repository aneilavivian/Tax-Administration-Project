# PANDUAN SETUP MACRO - TAX ADMINISTRATION SYSTEM

## ⚠️ PENTING: Sebelum Memulai
Pastikan Anda menggunakan **Excel Desktop** (bukan Excel Online), karena VBA Macro hanya berjalan di Excel Desktop.

---

## STEP 1: DOWNLOAD FILE EXCEL

1. Download file `Tax-Administration-Template.xlsm` dari repository ini
2. Simpan di folder yang mudah diakses (Desktop atau Documents)
3. **JANGAN** ubah nama file atau extension (.xlsm)

---

## STEP 2: AKTIFKAN DEVELOPER TAB

Jika Developer tab belum muncul di Excel:

### **Untuk Windows:**
1. Buka Excel
2. Klik **File** → **Options**
3. Pilih **Customize Ribbon**
4. Di bagian kanan, cek box **Developer**
5. Klik **OK**

### **Untuk Mac:**
1. Buka Excel
2. Klik **Excel** → **Preferences**
3. Klik **Ribbon & Toolbar**
4. Cek box **Developer**
5. Klik **Save**

---

## STEP 3: BUKA FILE EXCEL & SETUP MACRO

### **3a. Buka File Excel**
1. Buka file `Tax-Administration-Template.xlsm`
2. Excel akan menampilkan **Security Warning**: "Macros have been disabled"
3. Klik **Enable Content** atau **Enable Macros**
   - Jika tidak muncul, lanjut ke step 3b

### **3b. Manual Setup (Jika Macro tidak otomatis enable)**
1. Klik menu **Developer**
2. Klik **Visual Basic** (atau tekan `Ctrl + Alt + F11`)
3. Di panel kiri, klik file `Tax-Administration-Template`
4. Klik kanan → **Insert** → **Module**
5. Sebuah module baru akan terbuat (Module1)
6. Copy-paste seluruh kode VBA dari file `Macro-Code-Complete.vba` ke module ini
7. Klik **File** → **Close and Return to Microsoft Excel**

---

## STEP 4: TEST MACRO

1. Buka sheet **"LOGIN"**
2. Klik tombol berwarna **"LOGIN"**
3. Jika macro berfungsi, akan muncul:
   - Popup untuk input username & password
   - **Default:** Username: `admin` | Password: `123456`
4. Setelah login berhasil, Anda akan otomatis pindah ke sheet **"DASHBOARD"**

---

## STEP 5: TROUBLESHOOTING

### **Masalah: "Macro tidak berjalan saat klik tombol"**
**Solusi:**
- Pastikan sudah klik "Enable Macros"
- Buka Developer → Visual Basic
- Pastikan ada kode di Module1
- Cek nama tombol (harus cocok dengan nama sub/macro)

### **Masalah: "Error 'Sub or Function not defined'"**
**Solusi:**
- Copy-paste seluruh kode macro dari `Macro-Code-Complete.vba`
- Pastikan tidak ada bagian yang terpotong

### **Masalah: "Excel freeze/tidak responsive"**
**Solusi:**
- Tekan `Ctrl + Break` untuk stop macro
- Restart Excel
- Cek logic di dalam macro

### **Masalah: "Dropdown NPWP tidak muncul saat input"**
**Solusi:**
- Pastikan data di sheet "DATA-WP" sudah ada
- Cek bahwa dropdown sudah dikonfigurasi di form input
- Refresh data: Menu → Data Validation

---

## STEP 6: KONFIGURASI SISTEM (OPTIONAL)

### **Ubah Username & Password:**
1. Buka sheet **"ADMIN"** (atau hidden sheet)
2. Cari tabel daftar user
3. Ubah username & password sesuai keinginan
4. Update macro `LoginUser()` jika perlu

### **Ubah Tarif Pajak:**
1. Buka sheet **"MASTER-TARIF"**
2. Edit tabel tarif sesuai ketentuan
3. Semua perhitungan pajak otomatis menggunakan tarif ini

### **Tambah/Ubah Wajib Pajak:**
1. Buka sheet **"MASTER-WP"**
2. Klik tombol **"TAMBAH"**
3. Form popup akan muncul
4. Isi data NPWP, Nama, Jenis, Alamat
5. Klik **"SIMPAN"**

---

## STEP 7: READY TO USE! 🎉

Sekarang sistem siap digunakan. Baca file **USER-MANUAL.md** untuk panduan lengkap cara menggunakan semua fitur.

---

## QUICK REFERENCE - Tombol/Macro Utama

| Tombol | Sheet | Fungsi |
|--------|-------|--------|
| LOGIN | LOGIN | Validasi user & masuk dashboard |
| MASTER DATA | DASHBOARD | Manage wajib pajak & tarif |
| INPUT PAJAK | DASHBOARD | Form input pajak (PPh OP, PPN, dll) |
| PERHITUNGAN | DASHBOARD | Lihat rekapitulasi perhitungan |
| PELAPORAN | DASHBOARD | Generate & print laporan |
| SPT | DASHBOARD | Generate SPT format DJP |
| LOGOUT | Dashboard/Menu | Keluar & kembali login |

---

## KONTAK SUPPORT

Jika ada error atau pertanyaan teknis:
1. Cek troubleshooting di atas
2. Baca dokumentasi di file `Macro-Code-Complete.vba`
3. Hubungi pembuat template untuk bantuan

**Semoga lancar! 🚀**
