# TerbilangID

**TerbilangID** adalah fungsi VBA (Visual Basic for Applications) untuk Microsoft Excel yang mengubah angka menjadi teks terbilang dalam Bahasa Indonesia secara akurat dan fleksibel. 

*The ultimate Indonesian number-to-words converter for Excel.*

---

## ✨ Fitur Utama (Key Features)

* **Infinite Scale**: Mendukung hingga 66 digit (Vigintiliun).
* **Currency Mode**: Otomatis menambahkan "Rupiah" dan "Sen".
* **Decimal Support**: Menangani angka di belakang koma dengan presisi.
* **Smart Cleaning**: Otomatis membersihkan input dari simbol 'Rp', titik ribuan, atau spasi.
* **3 Modes**: Standar (Fisika/Matematika), Mata Uang (Kwitansi), dan Eja Digit (Nomor HP/ID).
* **Proper Case**: Output otomatis menggunakan huruf besar di setiap awal kata.

---

## 🚀 Cara Penggunaan (How to Use)

### 1. Instalasi (Installation)
1. Buka Microsoft Excel.
2. Tekan `ALT + F11` untuk membuka VBA Editor.
3. Klik `Insert` > `Module`.
4. **Copy-Paste** kode dari file `TerbilangMaster.bas` (atau kode di bawah) ke dalam modul tersebut.
5. Simpan file Excel Anda sebagai **Excel Macro-Enabled Workbook (.xlsm)**.

### 2. Rumus di Excel (Excel Formula)
Gunakan rumus berikut di sel Excel Anda:

| Mode | Deskripsi (Description) | Contoh Rumus (Example) |
| :--- | :--- | :--- |
| **0** | **Standar** (Standard) | `=TerbilangMaster(A1; 0)` |
| **1** | **Mata Uang** (Currency) | `=TerbilangMaster(A1; 1)` |
| **2** | **Eja Digit** (Digit Spelling) | `=TerbilangMaster(A1; 2)` |

---

## 📖 Sintaks (Syntax)
`=TerbilangMaster(Angka; [Mode]; [TulisSaja]; [Satuan]; [Sen])`

* **Angka**: Sel atau nilai angka.
* **Mode**: 0 = Standar, 1 = Rupiah, 2 = Eja.
* **TulisSaja**: (True/False) Menambahkan kata "Saja" di akhir (khusus Mode 1).
* **Satuan**: Custom satuan (Default: "Rupiah").
* **Sen**: Custom satuan desimal (Default: "Sen").

---

## 🛠 Kode Sumber (Source Code)



```vba
Option Explicit

' ==============================================================================
' FUNGSI: TerbilangId (The Indonesian Number-to-Word Converter)
' FITUR:
'   - Mendukung 66 Digit (skala Vigintiliun)
'   - Penanganan Minus/Negatif otomatis
'   - 3 Mode: [0] Standar/Fisika, [1] Mata Uang (Rupiah), [2] Eja Digit (Nomor HP)
'   - Input Cleaning (Membersihkan Rp, Titik, Spasi secara otomatis)
' ==============================================================================

Function TerbilangId(ByVal RefInput As Variant, _
                         Optional ByVal Mode As Integer = 0, _
                         Optional ByVal TulisSaja As Boolean = True, _
                         Optional ByVal CustomSatuan As String = "Rupiah", _
                         Optional ByVal CustomSen As String = "Sen") As String
    
    Dim strClean As String
    Dim bagianBulat As String, bagianDesimal As String
    Dim hasilBulat As String, hasilDesimal As String
    Dim IsMinus As Boolean
    Dim posKoma As Integer
    
    strClean = CStr(RefInput)
    strClean = Trim(Replace(strClean, "Rp", "")) ' Hapus simbol mata uang
    strClean = Replace(strClean, " ", "")        ' Hapus spasi
    
    ' Deteksi Minus
    If Left(strClean, 1) = "-" Or Left(strClean, 1) = "(" Then
        IsMinus = True
        strClean = Replace(Replace(strClean, "-", ""), "(", "")
        strClean = Replace(strClean, ")", "")
    End If
    
    ' Standarisasi Desimal (Ubah Titik Ribuan jadi kosong, Koma jadi Titik Sistem)
    ' Asumsi: Format Indonesia (1.000,00)
    strClean = Replace(strClean, ".", "")
    strClean = Replace(strClean, ",", ".")
    
    ' Validasi Angka
    If Not IsNumeric(strClean) Then
        TerbilangId = "#VALUE!" ' Kembalikan error Excel jika bukan angka
        Exit Function
    End If
    
    ' Pisahkan Bulat dan Desimal
    posKoma = InStr(strClean, ".")
    If posKoma > 0 Then
        bagianBulat = Left(strClean, posKoma - 1)
        bagianDesimal = Mid(strClean, posKoma + 1)
    Else
        bagianBulat = strClean
        bagianDesimal = ""
    End If
    
    ' --- LOGIKA UTAMA BERDASARKAN MODE ---
    Select Case Mode
        
        Case 2 ' --- MODE EJA DIGIT (Misal: No HP / Kode) ---
            hasilBulat = Core_EjaDigit(bagianBulat)
            If bagianDesimal <> "" Then
                hasilDesimal = " Koma " & Core_EjaDigit(bagianDesimal)
            End If
            
        Case 1 ' --- MODE MATA UANG (Keuangan) ---
            ' Proses Angka Bulat
            If val(bagianBulat) = 0 Then
                hasilBulat = "Nol " & CustomSatuan
            Else
                hasilBulat = Core_AngkaBesar(bagianBulat) & " " & CustomSatuan
            End If
            
            ' Proses Sen (Ambil 2 digit pertama saja untuk uang)
            If bagianDesimal <> "" Then
                Dim nSen As String
                nSen = Left(bagianDesimal & "00", 2) ' Pastikan minimal 2 digit
                If val(nSen) > 0 Then
                    hasilDesimal = " " & Core_Ratusan(CInt(nSen)) & " " & CustomSen
                ElseIf TulisSaja Then
                    hasilDesimal = " Saja"
                End If
            ElseIf TulisSaja Then
                hasilDesimal = " Saja"
            End If
            
        Case Else ' --- MODE 0: STANDAR (Matematika/Fisika) ---
            If val(bagianBulat) = 0 Then hasilBulat = "Nol" Else hasilBulat = Core_AngkaBesar(bagianBulat)
            
            If bagianDesimal <> "" Then
                hasilDesimal = " Koma " & Core_EjaDigit(bagianDesimal)
            End If
            
    End Select
    
    ' --- FINALISASI ---
    Dim FinalText As String
    FinalText = Trim(hasilBulat & hasilDesimal)
    
    If IsMinus And FinalText <> "Nol" Then
        FinalText = "Minus " & FinalText
    End If
    
    ' Huruf Besar di Awal Kata
    TerbilangId = Application.Proper(FinalText)
    
End Function

' [INTI 1] Menangani Angka Raksasa (Ribuan, Jutaan, ... Vigintiliun)
Private Function Core_AngkaBesar(ByVal strAngka As String) As String
    Dim Skala As Variant
    Dim nBlok As Integer, i As Integer
    Dim blok3 As String, hasilBlok As String, strHasil As String
    
    ' Definisi Skala hingga 10^63
    Skala = Array("", "Ribu", "Juta", "Milyar", "Triliun", "Kuadriliun", "Kuintiliun", "Sekstiliun", _
                  "Septiliun", "Oktiliun", "Noniliun", "Desiliun", "Undesiliun", "Duodesiliun", _
                  "Tredesiliun", "Kuatuordesiliun", "Kuindesiliun", "Seksdesiliun", "Septendesiliun", _
                  "Oktodesiliun", "Novemdesiliun", "Vigintiliun")
                  
    ' Normalisasi panjang string agar kelipatan 3 (Padding Zero)
    Do While Len(strAngka) Mod 3 <> 0
        strAngka = "0" & strAngka
    Loop
    
    nBlok = Len(strAngka) / 3
    
    For i = 1 To nBlok
        ' Ambil 3 digit dari kiri
        blok3 = Mid(strAngka, (i - 1) * 3 + 1, 3)
        hasilBlok = Core_Ratusan(CInt(blok3))
        
        If hasilBlok <> "" Then
            ' Logika Khusus: "Seribu" (Bukan Satu Ribu)
            ' Syarat: Bloknya adalah Ribuan, nilainya 1, dan blok sebelumnya kosong (misal 1.000, bukan 1.001.000)
            If CInt(blok3) = 1 And Skala(nBlok - i) = "Ribu" And (nBlok - i + 1) = nBlok Then
                strHasil = strHasil & " Seribu"
            Else
                strHasil = strHasil & " " & hasilBlok & " " & Skala(nBlok - i)
            End If
        End If
    Next i
    
    Core_AngkaBesar = Trim(strHasil)
End Function

' [INTI 2] Menangani 0 - 999
Private Function Core_Ratusan(ByVal n As Integer) As String
    Dim Satuan As Variant
    Satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
    
    Select Case n
        Case 0: Core_Ratusan = ""
        Case 1 To 11: Core_Ratusan = Satuan(n)
        Case 12 To 19: Core_Ratusan = Satuan(n Mod 10) & " Belas"
        Case 20 To 99: Core_Ratusan = Satuan(n \ 10) & " Puluh " & Satuan(n Mod 10)
        Case 100 To 199: Core_Ratusan = "Seratus " & Core_Ratusan(n Mod 100)
        Case 200 To 999: Core_Ratusan = Satuan(n \ 100) & " Ratus " & Core_Ratusan(n Mod 100)
    End Select
End Function

' [INTI 3] Mengeja Digit Satu per Satu (Untuk Desimal Teknis / No HP)
Private Function Core_EjaDigit(ByVal strAngka As String) As String
    Dim i As Integer, Char As String, Hasil As String
    For i = 1 To Len(strAngka)
        Char = Mid(strAngka, i, 1)
        Select Case Char
            Case "0": Hasil = Hasil & " Nol"
            Case "1": Hasil = Hasil & " Satu"
            Case "2": Hasil = Hasil & " Dua"
            Case "3": Hasil = Hasil & " Tiga"
            Case "4": Hasil = Hasil & " Empat"
            Case "5": Hasil = Hasil & " Lima"
            Case "6": Hasil = Hasil & " Enam"
            Case "7": Hasil = Hasil & " Tujuh"
            Case "8": Hasil = Hasil & " Delapan"
            Case "9": Hasil = Hasil & " Sembilan"
        End Select
    Next i
    Core_EjaDigit = Trim(Hasil)
End Function

' ==============================================================================
' FUNGSI BANTUAN (CHEAT SHEET)
' Cara Pakai: Ketik =TerbilangHelp() di sel mana saja.
' ==============================================================================

Function TerbilangHelp() As String
    Dim Msg As String
    Dim NL As String
    NL = Chr(10) ' Karakter Enter/Baris Baru untuk Sel Excel
    
    Msg = "=== PANDUAN TERBILANG MASTER ===" & NL & NL
    
    Msg = Msg & "[SINTAKS]" & NL
    Msg = Msg & "=TerbilangId(RefInput; [Mode]; [TulisSaja]; [Satuan]; [Sen])" & NL & NL
    
    Msg = Msg & "[PILIHAN MODE]" & NL
    Msg = Msg & "0 = Standar / Desimal (Contoh: Seratus Koma Lima)" & NL
    Msg = Msg & "1 = Mata Uang (Contoh: Seratus Rupiah Lima Puluh Sen)" & NL
    Msg = Msg & "2 = Eja Digit (Contoh: Nol Delapan Satu Dua...)" & NL & NL
    
    Msg = Msg & "[OPSIONAL]" & NL
    Msg = Msg & "TulisSaja : TRUE (Default) untuk tambah kata 'Saja' jika tidak ada sen." & NL
    Msg = Msg & "CustomSatuan : Ganti 'Rupiah' dengan kata lain (misal: 'Dolar')." & NL & NL
    
    Msg = Msg & "Catatan: Pastikan Format Cell diatur ke 'Wrap Text' agar tulisan ini rapi."
    
    TerbilangHelp = Msg
End Function

' ==============================================================================
' FUNGSI EKSTRA: DESKRIPSI RUMUS (Agar muncul keterangan saat ketik rumus)
' Cara Pakai: Jalankan Makro ini SEKALI SAJA (Klik Run / F5).
' ==============================================================================

Sub RegisterTerbilangDescription()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Args() As Variant
    Dim ArgDesc() As Variant
    
    FuncName = "TerbilangId"
    FuncDesc = "Mengubah angka menjadi huruf (Terbilang). Support hingga 66 digit, desimal, dan minus."
    
    Args = Array("RefInput", "Mode", "TulisSaja", "CustomSatuan", "CustomSen")
    ArgDesc = Array("Angka atau Sel yang ingin diubah.", _
                    "[0]=Standar, [1]=Rupiah/Uang, [2]=Eja Digit.", _
                    "TRUE = Tambah kata 'Saja' di akhir (Default). FALSE = Hapus.", _
                    "Ganti kata 'Rupiah' (jika Mode=1).", _
                    "Ganti kata 'Sen' (jika Mode=1).")
    
    Application.MacroOptions Macro:=FuncName, _
                             Description:=FuncDesc, _
                             ArgumentDescriptions:=ArgDesc, _
                             Category:=7 ' Kategori Text
                             
    MsgBox "Deskripsi Rumus TerbilangId berhasil ditambahkan ke Excel!", vbInformation, "Sukses"
End Sub
