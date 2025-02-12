Sub CopyDataToDailyDatabase()
    ' Deklarasi atribut
    Dim wbInput As Workbook
    Dim wbDatabase As Workbook
    Dim wsInput As Worksheet
    Dim wsDatabase As Worksheet
    Dim sourceRange1 As Range, sourceRange2 As Range
    Dim targetRow1 As Long, targetRow2 As Long
    Dim targetPath As String
    Dim deleteRange As Range

    ' Path file daily_database
    targetPath = "C:\Users\michael\Documents\daily_database.xlsx"

    ' Buka file daily_database jika belum terbuka [bag. Debugging]
    If wbDatabase Is Nothing Then
        On Error Resume Next
        Set wbDatabase = Workbooks.Open(targetPath)
        On Error GoTo 0
        If wbDatabase Is Nothing Then
            MsgBox "Gagal membuka file daily_database.xlsx.", vbCritical
            Exit Sub
        End If
    End If

    ' Set worksheet
    Set wbInput = ThisWorkbook
    Set wsInput = wbInput.Sheets("Input")
    Set wsDatabase = wbDatabase.Sheets("Daily Database")

    ' Validasi worksheet [bag. Debugging]
    If wsInput Is Nothing Or wsDatabase Is Nothing Then
        MsgBox "Sheet 'Input' atau 'Daily Database' tidak ditemukan.", vbCritical
        Exit Sub
    End If
    
    ' Ambil baris terakhir di sheet Daily Database
    If Application.WorksheetFunction.CountA(wsDatabase.Rows(4)) = 0 Then
        ' Jika sheet kosong (baris 4 kosong), mulai dari baris ke-4
        dbLastRow = 3
    Else
        ' Cari baris terakhir
        dbLastRow = wsDatabase.Cells(wsDatabase.Rows.Count, 1).End(xlUp).Row
    End If
    
    ' Ambil baris data dari sheet Input (data awal yang akan dicopy)
    Set inputRow = wsInput.Range("C3:G3") ' Data: Date, Shift, Machine, Size
    
    ' Ambil cell hasil pengecekan
    Set checkResults = wsInput.Range("C15:E15") ' Data hasil pengecekan
    
    ' Hitung jumlah cell yang terisi di range hasil pengecekan
    numFilledChecks = 0
    For Each cell In checkResults
        If Not IsEmpty(cell.Value) Then
            numFilledChecks = numFilledChecks + 1
        End If
    Next cell
    
    ' Salin data awal sebanyak jumlah cell hasil pengecekan yang terisi
    For i = 1 To numFilledChecks
        wsDatabase.Cells(dbLastRow + i, 1).Resize(, inputRow.Columns.Count).Value = inputRow.Value
    Next i
    
    ' Jika tidak ada pengecekan yang dilakukan, beri pesan dan hentikan
    If numFilledChecks = 0 Then
        wsInput.Activate
        Range("C3").Select
        MsgBox "Tidak ada hasil pengecekan yang terisi. Data tidak akan dicopy.", vbExclamation
        Exit Sub
    End If

    ' Set range sumber data pertama (C6:E34) (clip ring - air press)
    On Error Resume Next
    Set sourceRange1 = Union( _
        wsInput.Range("C6,D6,E6"), wsInput.Range("C7,D7,E7"), wsInput.Range("C9,D9,E9"), _
        wsInput.Range("C10,D10,E10"), wsInput.Range("C12,D12,E12"), wsInput.Range("C13,D13,E13"), _
        wsInput.Range("C15,D15,E15"), wsInput.Range("C16,D16,E16"), wsInput.Range("C18,D18,E18"), _
        wsInput.Range("C20,D20,E20"), wsInput.Range("C21,D21,E21"), wsInput.Range("C23,D23,E23"), _
        wsInput.Range("C24,D24,E24"), wsInput.Range("C26,D26,E26"), wsInput.Range("C27,D27,E27"), _
        wsInput.Range("C29,D29,E29"), wsInput.Range("C30,D30,E30"), wsInput.Range("C31,D31,E31"), _
        wsInput.Range("C32,D32,E32"), wsInput.Range("C34,D34,E34") _
    )
    On Error GoTo 0
    If sourceRange1 Is Nothing Then
        MsgBox "Source range 1 kosong atau tidak valid.", vbCritical
        Exit Sub
    End If

    ' Set range sumber data kedua (H6:J18) (green tire - keterangan)
    On Error Resume Next
    Set sourceRange2 = Union( _
        wsInput.Range("H6,I6,J6"), wsInput.Range("H8,I8,J8"), wsInput.Range("H9,I9,J9"), _
        wsInput.Range("H10,I10,J10"), wsInput.Range("H11,I11,J11"), wsInput.Range("H12,I12,J12"), _
        wsInput.Range("H14,I14,J14"), wsInput.Range("H15,I15,J15"), wsInput.Range("H16,I16,J16"), _
        wsInput.Range("H17,I17,J17"), wsInput.Range("H18,I18,J18") _
    )
    On Error GoTo 0
    If sourceRange2 Is Nothing Then
        MsgBox "Source range 2 kosong atau tidak valid.", vbCritical
        Exit Sub
    End If

    ' Tentukan baris berikutnya untuk data pertama di Daily Database, mulai dari F4
    If Application.WorksheetFunction.CountA(wsDatabase.Columns("F")) = 0 Then
        targetRow1 = 4 ' Baris pertama untuk data pertama
    Else
        targetRow1 = wsDatabase.Cells(wsDatabase.Rows.Count, "F").End(xlUp).Row + 1
    End If

    ' Salin data pertama secara transpose ke Daily Database
    sourceRange1.Copy
    wsDatabase.Cells(targetRow1, "F").PasteSpecial Paste:=xlPasteValues, Transpose:=True

    ' Tentukan baris berikutnya untuk data kedua, mulai dari kolom Z
    If Application.WorksheetFunction.CountA(wsDatabase.Columns("Z")) = 0 Then
        targetRow2 = 4 ' Baris pertama untuk data kedua
    Else
        targetRow2 = wsDatabase.Cells(wsDatabase.Rows.Count, "Z").End(xlUp).Row + 1
    End If

    ' Salin data kedua secara transpose ke Daily Database
    sourceRange2.Copy
    wsDatabase.Cells(targetRow2, "Z").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False

    ' Bersihkan data di sheet Input
    wsInput.Activate
    Set deleteRange = Union( _
        wsInput.Range("H6,H8,H9,I9,J9,H10,I10,J10,H11,I11,J11,H12,I12,J12,H14,H15,I15,J15,H16,I16,J16,H17,I17,J17,H18,I18,J18"), _
        wsInput.Range("C6,C7,C9,C10,C12,D12,E12,C13,D13,E13,C15,D15,E15,C16,D16,E16"), _
        wsInput.Range("C18,D18,E18,C20,C21,C23,D23,E23,C24,D24,E24,C26,D26,E26,C27,D27,E27"), _
        wsInput.Range("C29,E29,C30,E30,C31,E31,C32,E32,C34,D3,E3,F3") _
    )
    deleteRange.ClearContents

    ' Simpan file daily_database dan buka wsInput
    wbDatabase.Save
    wbDatabase.Close
    wsInput.Activate
    Range("C3").Select

    ' Pesan sukses
    MsgBox "Data berhasil disalin ke Daily Database!", vbInformation
End Sub
