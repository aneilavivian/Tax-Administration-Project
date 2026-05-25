' =====================================================
' TAX ADMINISTRATION SYSTEM - COMPLETE VBA MACRO CODE
' =====================================================
' Author: Tax Administration Team
' Version: 1.0
' Last Updated: 2026
' 
' INSTRUCTIONS:
' 1. Copy seluruh kode ini
' 2. Buka Excel → Developer → Visual Basic (Ctrl+Alt+F11)
' 3. Di panel kiri, klik file Excel Anda
' 4. Insert → Module
' 5. Paste kode di module baru
' 6. Tutup VBA Editor
' =====================================================

' ==================== MODULE 1: LOGIN SYSTEM ====================

Sub LoginUser()
'
' Macro untuk validasi login user
' Username default: admin
' Password default: 123456
'

Dim userName As String
Dim userPass As String
Dim ws As Worksheet

' Buka sheet LOGIN
Set ws = ThisWorkbook.Sheets("LOGIN")
ws.Select

' Input username & password
userName = InputBox("Masukkan Username:", "Login System")

If userName = "" Then
    MsgBox "Login dibatalkan!", vbInformation, "Tax Administration"
    Exit Sub
End If

userPass = InputBox("Masukkan Password:", "Login System")

If userPass = "" Then
    MsgBox "Login dibatalkan!", vbInformation, "Tax Administration"
    Exit Sub
End If

' Validasi user
If userName = "admin" And userPass = "123456" Then
    MsgBox "Selamat datang, " & userName & "!", vbInformation, "Login Berhasil"
    
    ' Pindah ke DASHBOARD
    ThisWorkbook.Sheets("DASHBOARD").Select
    
    ' Optional: Clear form
    Call ClearForm
    
Else
    MsgBox "Username atau Password salah! Silakan coba lagi.", vbCritical, "Login Gagal"
    Call LoginUser ' Repeat login
End If

End Sub

Sub Logout()
'
' Macro untuk logout dan kembali ke login
'

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("LOGIN")

' Clear semua sheet
Call ClearAllSheets

' Pindah ke sheet LOGIN
ws.Select

MsgBox "Anda telah logout. Silakan login kembali.", vbInformation, "Logout"

End Sub

Sub ClearAllSheets()
'
' Kosongkan semua form input
'

Dim ws As Worksheet

' Clear sheet-sheet input
For Each ws In ThisWorkbook.Sheets
    If ws.Name Like "INPUT-*" Or ws.Name Like "TEMP-*" Then
        ws.UsedRange.Clear
    End If
Next ws

End Sub

Sub ClearForm()
'
' Kosongkan form input saat ini
'

Dim ws As Worksheet
Set ws = ActiveSheet

' Clear only form cells (A1:F50)
ws.Range("A1:F50").Clear

End Sub


' ==================== MODULE 2: MASTER DATA MANAGEMENT ====================

Sub AddWajibPajak()
'
' Macro untuk menambah data wajib pajak baru
'

Dim ws As Worksheet
Dim lastRow As Long
Dim npwp As String
Dim nama As String
Dim jenis As String
Dim alamat As String
Dim status As String

Set ws = ThisWorkbook.Sheets("DATA-WP")

' Cari baris terakhir
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

' Input data via dialog
npwp = InputBox("Masukkan NPWP (format: XX.XXX.XXX.X-XXX.XXX):", "Tambah Wajib Pajak")
If npwp = "" Then Exit Sub

nama = InputBox("Masukkan Nama Wajib Pajak:", "Tambah Wajib Pajak")
If nama = "" Then Exit Sub

jenis = InputBox("Masukkan Jenis (OP/Badan):", "Tambah Wajib Pajak")
If jenis = "" Then Exit Sub

alamat = InputBox("Masukkan Alamat:", "Tambah Wajib Pajak")
If alamat = "" Then Exit Sub

status = InputBox("Masukkan Status (Aktif/Tidak Aktif):", "Tambah Wajib Pajak")
If status = "" Then Exit Sub

' Validasi NPWP (simple check)
If Len(npwp) < 5 Then
    MsgBox "NPWP tidak valid!", vbCritical, "Error"
    Exit Sub
End If

' Input ke sheet
ws.Cells(lastRow, 1).Value = npwp
ws.Cells(lastRow, 2).Value = nama
ws.Cells(lastRow, 3).Value = jenis
ws.Cells(lastRow, 4).Value = alamat
ws.Cells(lastRow, 5).Value = status
ws.Cells(lastRow, 6).Value = Now() ' Tanggal input

MsgBox "Data Wajib Pajak berhasil ditambahkan!", vbInformation, "Success"

' Format sheet
Call FormatDataSheet

End Sub

Sub EditWajibPajak()
'
' Macro untuk edit data wajib pajak
'

Dim ws As Worksheet
Dim searchNPWP As String
Dim foundRow As Long
Dim i As Long

Set ws = ThisWorkbook.Sheets("DATA-WP")

searchNPWP = InputBox("Masukkan NPWP yang ingin diedit:", "Edit Wajib Pajak")
If searchNPWP = "" Then Exit Sub

' Cari NPWP
foundRow = 0
For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i, 1).Value = searchNPWP Then
        foundRow = i
        Exit For
    End If
Next i

If foundRow = 0 Then
    MsgBox "NPWP tidak ditemukan!", vbCritical, "Not Found"
    Exit Sub
End If

' Edit data
Dim newNama, newJenis, newAlamat, newStatus As String

newNama = InputBox("Edit Nama (tekan Enter jika tidak diubah):", , ws.Cells(foundRow, 2).Value)
If newNama <> "" Then ws.Cells(foundRow, 2).Value = newNama

newJenis = InputBox("Edit Jenis OP/Badan (tekan Enter jika tidak diubah):", , ws.Cells(foundRow, 3).Value)
If newJenis <> "" Then ws.Cells(foundRow, 3).Value = newJenis

newAlamat = InputBox("Edit Alamat (tekan Enter jika tidak diubah):", , ws.Cells(foundRow, 4).Value)
If newAlamat <> "" Then ws.Cells(foundRow, 4).Value = newAlamat

newStatus = InputBox("Edit Status (tekan Enter jika tidak diubah):", , ws.Cells(foundRow, 5).Value)
If newStatus <> "" Then ws.Cells(foundRow, 5).Value = newStatus

ws.Cells(foundRow, 6).Value = Now() ' Update tanggal perubahan

MsgBox "Data Wajib Pajak berhasil diubah!", vbInformation, "Success"

End Sub

Sub DeleteWajibPajak()
'
' Macro untuk hapus data wajib pajak
'

Dim ws As Worksheet
Dim searchNPWP As String
Dim foundRow As Long
Dim i As Long
Dim confirm As VbMsgBoxResult

Set ws = ThisWorkbook.Sheets("DATA-WP")

searchNPWP = InputBox("Masukkan NPWP yang ingin dihapus:", "Hapus Wajib Pajak")
If searchNPWP = "" Then Exit Sub

' Cari NPWP
foundRow = 0
For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i, 1).Value = searchNPWP Then
        foundRow = i
        Exit For
    End If
Next i

If foundRow = 0 Then
    MsgBox "NPWP tidak ditemukan!", vbCritical, "Not Found"
    Exit Sub
End If

' Konfirmasi hapus
confirm = MsgBox("Yakin ingin menghapus data ini?", vbYesNo + vbQuestion, "Konfirmasi Hapus")

If confirm = vbYes Then
    ws.Rows(foundRow).Delete
    MsgBox "Data berhasil dihapus!", vbInformation, "Success"
Else
    MsgBox "Penghapusan dibatalkan.", vbInformation, "Cancelled"
End If

End Sub

Sub SearchWajibPajak()
'
' Macro untuk mencari wajib pajak
'

Dim ws As Worksheet
Dim searchNPWP As String
Dim i As Long
Dim found As Boolean

Set ws = ThisWorkbook.Sheets("DATA-WP")

searchNPWP = InputBox("Masukkan NPWP atau Nama yang dicari:", "Cari Wajib Pajak")
If searchNPWP = "" Then Exit Sub

found = False

For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i, 1).Value = searchNPWP Or ws.Cells(i, 2).Value Like "*" & searchNPWP & "*" Then
        MsgBox "Ditemukan:" & vbCrLf & _
               "NPWP: " & ws.Cells(i, 1).Value & vbCrLf & _
               "Nama: " & ws.Cells(i, 2).Value & vbCrLf & _
               "Jenis: " & ws.Cells(i, 3).Value & vbCrLf & _
               "Alamat: " & ws.Cells(i, 4).Value & vbCrLf & _
               "Status: " & ws.Cells(i, 5).Value, vbInformation, "Hasil Pencarian"
        
        ws.Rows(i).Select
        found = True
        Exit For
    End If
Next i

If Not found Then
    MsgBox "Data tidak ditemukan!", vbInformation, "Not Found"
End If

End Sub


' ==================== MODULE 3: TAX CALCULATION ====================

Function CalculatePPhOP(penghasilanBruto As Double, biaya As Double) As Double
'
' Hitung PPh OP berdasarkan bracket tarif
' Asumsi: Tarif progresif Indonesia
'

Dim penghasilanNeto As Double
Dim pph As Double

' PKP (Penghasilan Kena Pajak) = Penghasilan Bruto - Biaya
penghasilanNeto = penghasilanBruto - biaya

' Bracket Tarif PPh OP (2024)
If penghasilanNeto <= 60000000 Then
    pph = penghasilanNeto * 0.05
ElseIf penghasilanNeto <= 250000000 Then
    pph = (60000000 * 0.05) + ((penghasilanNeto - 60000000) * 0.15)
ElseIf penghasilanNeto <= 500000000 Then
    pph = (60000000 * 0.05) + (190000000 * 0.15) + ((penghasilanNeto - 250000000) * 0.25)
Else
    pph = (60000000 * 0.05) + (190000000 * 0.15) + (250000000 * 0.25) + ((penghasilanNeto - 500000000) * 0.3)
End If

CalculatePPhOP = pph

End Function

Function CalculatePPhBadan(penghasilanNeto As Double) As Double
'
' Hitung PPh Badan (Tarif 20% untuk bendel profit 2024)
'

Dim pph As Double

pph = penghasilanNeto * 0.20 ' Tarif standar PPh Badan

CalculatePPhBadan = pph

End Function

Function CalculatePPN(dpp As Double, rate As Double) As Double
'
' Hitung PPN
' DPP = Dasar Pengenaan Pajak
' Rate = Tarif PPN (default 11%)
'

Dim ppn As Double

ppn = dpp * (rate / 100)

CalculatePPN = ppn

End Function

Function CalculatePPh21(gaji As Double) As Double
'
' Hitung PPh Pasal 21 (Gaji Karyawan)
' Simplified calculation
'

Dim pph21 As Double
Dim pkp As Double

' Asumsi: PTKP 2024 = Rp 54.000.000
Dim ptkp As Double
ptkp = 54000000

' PKP = Gaji - PTKP
pkp = gaji - ptkp

If pkp <= 0 Then
    pph21 = 0
Else
    ' Tarif progresif seperti PPh OP
    If pkp <= 60000000 Then
        pph21 = pkp * 0.05
    ElseIf pkp <= 250000000 Then
        pph21 = (60000000 * 0.05) + ((pkp - 60000000) * 0.15)
    Else
        pph21 = (60000000 * 0.05) + (190000000 * 0.15) + ((pkp - 250000000) * 0.25)
    End If
End If

CalculatePPh21 = pph21

End Function

Function CalculatePPh23(dasar As Double, tipe As String) As Double
'
' Hitung PPh Pasal 23
' Tipe: "Dividen", "Bunga", "Royalti", "Jasa"
'

Dim pph23 As Double
Dim tarif As Double

Select Case tipe
    Case "Dividen"
        tarif = 0.10 ' 10%
    Case "Bunga"
        tarif = 0.15 ' 15%
    Case "Royalti"
        tarif = 0.15 ' 15%
    Case "Jasa"
        tarif = 0.02 ' 2%
    Case Else
        tarif = 0.10 ' Default
End Select

pph23 = dasar * tarif

CalculatePPh23 = pph23

End Function


' ==================== MODULE 4: INPUT & REPORTING ====================

Sub InputDataPPh_OP()
'
' Macro untuk input data PPh OP
'

Dim ws As Worksheet
Dim lastRow As Long
Dim npwp As String
Dim nama As String
Dim penghasilanBruto As Double
Dim biaya As Double
Dim pph As Double

Set ws = ThisWorkbook.Sheets("DATA-PPH-OP")

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

' Input via dialog
npwp = InputBox("Masukkan NPWP:", "Input PPh OP")
If npwp = "" Then Exit Sub

nama = InputBox("Masukkan Nama WP:", "Input PPh OP")
If nama = "" Then Exit Sub

penghasilanBruto = CDbl(InputBox("Masukkan Penghasilan Bruto:", "Input PPh OP"))
If penghasilanBruto = 0 Then Exit Sub

biaya = CDbl(InputBox("Masukkan Biaya (pengurang):", "Input PPh OP"))

' Hitung PPh
pph = CalculatePPhOP(penghasilanBruto, biaya)

' Simpan data
ws.Cells(lastRow, 1).Value = npwp
ws.Cells(lastRow, 2).Value = nama
ws.Cells(lastRow, 3).Value = penghasilanBruto
ws.Cells(lastRow, 4).Value = biaya
ws.Cells(lastRow, 5).Value = penghasilanBruto - biaya ' PKP
ws.Cells(lastRow, 6).Value = pph ' PPh OP
ws.Cells(lastRow, 7).Value = Now() ' Tanggal

MsgBox "Data PPh OP berhasil disimpan!" & vbCrLf & _
       "PPh OP: Rp " & Format(pph, "#,##0"), vbInformation, "Success"

End Sub

Sub InputDataPPN()
'
' Macro untuk input data PPN
'

Dim ws As Worksheet
Dim lastRow As Long
Dim npwp As String
Dim tanggal As Date
Dim dpp As Double
Dim rate As Double
Dim ppn As Double

Set ws = ThisWorkbook.Sheets("DATA-PPN")

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

' Input
npwp = InputBox("Masukkan NPWP Penjual:", "Input PPN")
If npwp = "" Then Exit Sub

dpp = CDbl(InputBox("Masukkan DPP (Dasar Pengenaan Pajak):", "Input PPN"))
If dpp = 0 Then Exit Sub

rate = CDbl(InputBox("Masukkan Rate PPN % (default 11):", , 11))

' Hitung PPN
ppn = CalculatePPN(dpp, rate)

' Simpan
ws.Cells(lastRow, 1).Value = npwp
ws.Cells(lastRow, 2).Value = Now() ' Tanggal
ws.Cells(lastRow, 3).Value = dpp
ws.Cells(lastRow, 4).Value = rate
ws.Cells(lastRow, 5).Value = ppn
ws.Cells(lastRow, 6).Value = "Belum Bayar" ' Status

MsgBox "Data PPN berhasil disimpan!" & vbCrLf & _
       "PPN: Rp " & Format(ppn, "#,##0"), vbInformation, "Success"

End Sub

Sub InputDataPPh21()
'
' Macro untuk input data PPh Pasal 21
'

Dim ws As Worksheet
Dim lastRow As Long
Dim npwpKaryawan As String
Dim nama As String
Dim gaji As Double
Dim tunjangan As Double
Dim potongan As Double
Dim pph21 As Double

Set ws = ThisWorkbook.Sheets("DATA-PPH21")

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

' Input
npwpKaryawan = InputBox("Masukkan NPWP Karyawan:", "Input PPh 21")
If npwpKaryawan = "" Then Exit Sub

nama = InputBox("Masukkan Nama Karyawan:", "Input PPh 21")
gaji = CDbl(InputBox("Masukkan Gaji Pokok:", "Input PPh 21"))
tunjangan = CDbl(InputBox("Masukkan Tunjangan:", "Input PPh 21", 0))
potongan = CDbl(InputBox("Masukkan Potongan (JKK, JKM, dll):", "Input PPh 21", 0))

' Hitung PPh 21
pph21 = CalculatePPh21(gaji + tunjangan - potongan)

' Simpan
ws.Cells(lastRow, 1).Value = npwpKaryawan
ws.Cells(lastRow, 2).Value = nama
ws.Cells(lastRow, 3).Value = gaji
ws.Cells(lastRow, 4).Value = tunjangan
ws.Cells(lastRow, 5).Value = potongan
ws.Cells(lastRow, 6).Value = gaji + tunjangan - potongan ' Penghasilan Neto
ws.Cells(lastRow, 7).Value = pph21
ws.Cells(lastRow, 8).Value = Now()

MsgBox "Data PPh 21 berhasil disimpan!" & vbCrLf & _
       "PPh 21: Rp " & Format(pph21, "#,##0"), vbInformation, "Success"

End Sub


' ==================== MODULE 5: REPORTING & SPT ====================

Sub GenerateReportPajakTertanggung()
'
' Macro untuk generate laporan pajak terutang
'

Dim ws As Worksheet
Dim wsSummary As Worksheet
Dim lastRow As Long

' Tentukan sheet untuk laporan
If Not SheetExists("REPORT-SUMMARY") Then
    Set wsSummary = ThisWorkbook.Sheets.Add
    wsSummary.Name = "REPORT-SUMMARY"
Else
    Set wsSummary = ThisWorkbook.Sheets("REPORT-SUMMARY")
    wsSummary.Cells.Clear
End If

' Header
wsSummary.Range("A1").Value = "LAPORAN PAJAK TERUTANG"
wsSummary.Range("A1").Font.Bold = True
wsSummary.Range("A1").Font.Size = 14

wsSummary.Range("A2").Value = "Tanggal: " & Format(Now(), "dd-mm-yyyy")

wsSummary.Range("A4").Value = "Jenis Pajak"
wsSummary.Range("B4").Value = "Jumlah (Rp)"
wsSummary.Range("A4").Font.Bold = True
wsSummary.Range("B4").Font.Bold = True

' Data PPh OP
Dim totalPPhOP As Double
Set ws = ThisWorkbook.Sheets("DATA-PPH-OP")
totalPPhOP = Application.WorksheetFunction.Sum(ws.Range("F:F"))

wsSummary.Range("A5").Value = "PPh OP"
wsSummary.Range("B5").Value = totalPPhOP
wsSummary.Range("B5").NumberFormat = "#,##0"

' Data PPN
Dim totalPPN As Double
Set ws = ThisWorkbook.Sheets("DATA-PPN")
totalPPN = Application.WorksheetFunction.Sum(ws.Range("E:E"))

wsSummary.Range("A6").Value = "PPN"
wsSummary.Range("B6").Value = totalPPN
wsSummary.Range("B6").NumberFormat = "#,##0"

' Data PPh 21
Dim totalPPh21 As Double
Set ws = ThisWorkbook.Sheets("DATA-PPH21")
totalPPh21 = Application.WorksheetFunction.Sum(ws.Range("G:G"))

wsSummary.Range("A7").Value = "PPh Pasal 21"
wsSummary.Range("B7").Value = totalPPh21
wsSummary.Range("B7").NumberFormat = "#,##0"

' Total
wsSummary.Range("A9").Value = "TOTAL PAJAK TERUTANG"
wsSummary.Range("A9").Font.Bold = True
wsSummary.Range("B9").Value = totalPPhOP + totalPPN + totalPPh21
wsSummary.Range("B9").Font.Bold = True
wsSummary.Range("B9").NumberFormat = "#,##0"

' Format column width
wsSummary.Columns("A").ColumnWidth = 25
wsSummary.Columns("B").ColumnWidth = 20

MsgBox "Laporan berhasil dibuat! Lihat sheet 'REPORT-SUMMARY'", vbInformation, "Success"

wsSummary.Select

End Sub

Sub GenerateSPT()
'
' Macro untuk generate SPT (Surat Pemberitahuan)
'

Dim ws As Worksheet
Dim wsSPT As Worksheet
Dim npwp As String

' Input NPWP
npwp = InputBox("Masukkan NPWP untuk membuat SPT:", "Generate SPT")
If npwp = "" Then Exit Sub

' Buat sheet SPT baru jika belum ada
If Not SheetExists("SPT-" & npwp) Then
    Set wsSPT = ThisWorkbook.Sheets.Add
    wsSPT.Name = "SPT-" & npwp
Else
    Set wsSPT = ThisWorkbook.Sheets("SPT-" & npwp)
    wsSPT.Cells.Clear
End If

' Header SPT
wsSPT.Range("A1").Value = "SURAT PEMBERITAHUAN (SPT)"
wsSPT.Range("A1").Font.Bold = True
wsSPT.Range("A1").Font.Size = 14

wsSPT.Range("A3").Value = "NPWP: " & npwp
wsSPT.Range("A4").Value = "Tanggal Laporan: " & Format(Now(), "dd-mm-yyyy")
wsSPT.Range("A5").Value = "Periode: " & Format(Now(), "mmmm yyyy")

' Detil SPT
wsSPT.Range("A7").Value = "PENGHASILAN"
wsSPT.Range("A7").Font.Bold = True

wsSPT.Range("A8").Value = "Penghasilan dari Usaha/Bisnis:"
wsSPT.Range("A9").Value = "Penghasilan dari Pekerjaan:"
wsSPT.Range("A10").Value = "Penghasilan dari Modal:"
wsSPT.Range("A11").Value = "Penghasilan Lainnya:"
wsSPT.Range("A13").Value = "Jumlah Penghasilan Bruto"
wsSPT.Range("A13").Font.Bold = True

wsSPT.Range("A15").Value = "PENGURANG"
wsSPT.Range("A15").Font.Bold = True

wsSPT.Range("A16").Value = "Biaya Operasional:"
wsSPT.Range("A17").Value = "Biaya Lainnya:"
wsSPT.Range("A19").Value = "Jumlah Pengurang"
wsSPT.Range("A19").Font.Bold = True

wsSPT.Range("A21").Value = "PENGHASILAN KENA PAJAK"
wsSPT.Range("A21").Font.Bold = True

wsSPT.Range("A23").Value = "PERHITUNGAN PAJAK"
wsSPT.Range("A23").Font.Bold = True

wsSPT.Range("A24").Value = "PPh Terutang:"
wsSPT.Range("A25").Value = "PPh Yang Sudah Dibayar:"
wsSPT.Range("A26").Value = "Kurang / Lebih Bayar:"
wsSPT.Range("A26").Font.Bold = True

' Format kolom
wsSPT.Columns("A").ColumnWidth = 30
wsSPT.Columns("B").ColumnWidth = 20

MsgBox "SPT untuk NPWP " & npwp & " berhasil dibuat!" & vbCrLf & _
       "Lihat sheet 'SPT-" & npwp & "'", vbInformation, "Success"

wsSPT.Select

End Sub


' ==================== MODULE 6: UTILITY FUNCTIONS ====================

Sub FormatDataSheet()
'
' Format sheet data dengan header
'

Dim ws As Worksheet
Set ws = ActiveSheet

' Tambahkan header jika belum ada
If ws.Range("A1").Value = "" Then
    ws.Range("A1").Value = "NPWP"
    ws.Range("B1").Value = "Nama"
    ws.Range("C1").Value = "Jenis"
    ws.Range("D1").Value = "Alamat"
    ws.Range("E1").Value = "Status"
    ws.Range("F1").Value = "Tanggal Input"
    
    ' Format header
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(192, 192, 192)
    
    ' Set column width
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 15
End If

End Sub

Function SheetExists(sheetName As String) As Boolean
'
' Check apakah sheet sudah ada
'

Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Sheets(sheetName)
SheetExists = (ws Is Nothing) = False
On Error GoTo 0

End Function

Sub PrintCurrentSheet()
'
' Macro untuk print sheet saat ini
'

ActiveSheet.PrintPreview

End Sub

Sub ExportToPDF()
'
' Macro untuk export ke PDF
'

Dim fileName As String
Dim filePath As String

fileName = InputBox("Masukkan nama file (tanpa .pdf):", "Export ke PDF")
If fileName = "" Then Exit Sub

' Simpan ke Desktop
filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & fileName & ".pdf"

' Export
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

MsgBox "File PDF berhasil dibuat di Desktop dengan nama: " & fileName & ".pdf", vbInformation, "Success"

End Sub


' ==================== MODULE 7: DASHBOARD NAVIGATION ====================

Sub OpenDashboard()
'
' Buka Dashboard
'

ThisWorkbook.Sheets("DASHBOARD").Select

End Sub

Sub OpenMasterData()
'
' Buka Master Data menu
'

Dim choice As String

choice = InputBox("Pilih opsi:" & vbCrLf & _
                  "1 = Lihat Daftar WP" & vbCrLf & _
                  "2 = Tambah WP" & vbCrLf & _
                  "3 = Edit WP" & vbCrLf & _
                  "4 = Hapus WP" & vbCrLf & _
                  "5 = Cari WP", "Master Data Menu")

Select Case choice
    Case "1"
        ThisWorkbook.Sheets("DATA-WP").Select
    Case "2"
        Call AddWajibPajak
    Case "3"
        Call EditWajibPajak
    Case "4"
        Call DeleteWajibPajak
    Case "5"
        Call SearchWajibPajak
    Case Else
        MsgBox "Pilihan tidak valid", vbCritical
End Select

End Sub

Sub OpenInputMenu()
'
' Buka menu input pajak
'

Dim choice As String

choice = InputBox("Pilih jenis pajak:" & vbCrLf & _
                  "1 = PPh OP" & vbCrLf & _
                  "2 = PPN" & vbCrLf & _
                  "3 = PPh Pasal 21" & vbCrLf & _
                  "4 = PPh Pasal 23" & vbCrLf & _
                  "5 = Kembali", "Input Pajak Menu")

Select Case choice
    Case "1"
        Call InputDataPPh_OP
    Case "2"
        Call InputDataPPN
    Case "3"
        Call InputDataPPh21
    Case "4"
        MsgBox "Fitur PPh 23 sedang dikembangkan", vbInformation
    Case "5"
        Call OpenDashboard
    Case Else
        MsgBox "Pilihan tidak valid", vbCritical
End Select

End Sub

Sub OpenPerhitungan()
'
' Buka menu perhitungan
'

Call GenerateReportPajakTertanggung

End Sub

Sub OpenSPT()
'
' Buka menu SPT
'

Call GenerateSPT

End Sub

' =====================================================
' END OF MACRO CODE
' =====================================================
' 
' SUMMARY MACRO YANG TERSEDIA:
' 
' LOGIN & AUTHENTICATION:
' - LoginUser() → Login dengan username/password
' - Logout() → Logout & kembali ke login
'
' MASTER DATA:
' - AddWajibPajak() → Tambah WP baru
' - EditWajibPajak() → Edit data WP
' - DeleteWajibPajak() → Hapus WP
' - SearchWajibPajak() → Cari WP
'
' CALCULATION:
' - CalculatePPhOP() → Hitung PPh OP
' - CalculatePPhBadan() → Hitung PPh Badan
' - CalculatePPN() → Hitung PPN
' - CalculatePPh21() → Hitung PPh 21
' - CalculatePPh23() → Hitung PPh 23
'
' INPUT DATA:
' - InputDataPPh_OP() → Input PPh OP
' - InputDataPPN() → Input PPN
' - InputDataPPh21() → Input PPh 21
'
' REPORTING:
' - GenerateReportPajakTertanggung() → Generate laporan
' - GenerateSPT() → Generate SPT
' - PrintCurrentSheet() → Print
' - ExportToPDF() → Export PDF
'
' NAVIGATION:
' - OpenDashboard() → Buka Dashboard
' - OpenMasterData() → Menu Master Data
' - OpenInputMenu() → Menu Input Pajak
' - OpenPerhitungan() → Buka Perhitungan
' - OpenSPT() → Buka SPT
'
' =====================================================
