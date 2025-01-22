Sub SaveWithConfirmation()
    ' Pesan konfirmasi sebelum menjalankan macro
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Periksa kembali data, apakah sudah yakin?", vbYesNo + vbQuestion, "Konfirmasi Save")
    
    If userResponse = vbYes Then
        Call CopyDataToDailyDatabase
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdInformationSheet_Click()
    Dim wbName As String
    Dim filePath As String

    ' Nama workbook dan path file
    wbName = "WorkbookTujuan.xlsx" ' Nama workbook
    filePath = "C:\Users\Username\Documents\WorkbookTujuan.xlsx" ' Path file

    On Error Resume Next
    ' Cek apakah workbook terbuka
    Workbooks(wbName).Activate
    If Err.Number <> 0 Then
        ' Jika workbook tidak terbuka, buka file
        If Dir(filePath) <> "" Then
            Workbooks.Open filePath
            MsgBox "File berhasil dibuka!", vbInformation
        Else
            MsgBox "File tidak ditemukan: " & filePath, vbExclamation
        End If
    Else
        MsgBox "Berpindah ke workbook: " & wbName, vbInformation
    End If
    On Error GoTo 0
End Sub
