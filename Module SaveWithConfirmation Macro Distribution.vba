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
