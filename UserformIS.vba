Private Sub UserForm_Initialize()

    Dim i As Integer

    txtDate.Value = Date

    ' Isi bulan filter
    For i = 1 To 12
        cmbFilterMonth.AddItem Format(DateSerial(Year(Date), i, 1), "mmmm")
    Next i

    cmbFilterStatus.AddItem "OPEN"
    cmbFilterStatus.AddItem "CLOSE"
    cmbFilterStatus.AddItem "ALL"

    cmbStatusM.AddItem "OPEN"
    cmbStatusM.AddItem "CLOSE"

    ' Setting ListBox
    lstEntry.ColumnCount = 5
    lstEntry.ColumnWidths = "80;100;120;150;50"

    lstManage.ColumnCount = 6
    lstManage.ColumnWidths = "0;80;100;120;50;70"

End Sub
=================================================================================================================================================================================

Private Sub btnAdd_Click()

    If Trim(txtParameter.Value) = "" Then
        MsgBox "Parameter kosong.", vbExclamation
        txtParameter.SetFocus
        Exit Sub
    End If

    If Trim(txtQty.Value) = "" Then
        MsgBox "Qty kosong.", vbExclamation
        txtQty.SetFocus
        Exit Sub
    End If

    lstEntry.AddItem txtDate.Value
    lstEntry.List(lstEntry.ListCount - 1, 1) = cmbSection.Value
    lstEntry.List(lstEntry.ListCount - 1, 2) = txtParameter.Value
    lstEntry.List(lstEntry.ListCount - 1, 3) = txtDesc.Value
    lstEntry.List(lstEntry.ListCount - 1, 4) = txtQty.Value

    txtParameter.Value = ""
    txtDesc.Value = ""
    txtQty.Value = ""

    txtParameter.SetFocus

End Sub
=================================================================================================================================================================================

Private Sub btnRemove_Click()

    If lstEntry.ListIndex = -1 Then Exit Sub
    lstEntry.RemoveItem lstEntry.ListIndex

End Sub
=================================================================================================================================================================================

Private Sub btnSaveAll_Click()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    If lstEntry.ListCount = 0 Then
        MsgBox "Tidak ada data untuk disimpan.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("NG_Database")

    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    For i = 0 To lstEntry.ListCount - 1

        ws.Cells(lastRow, 1).Value = lstEntry.List(i, 0)
        ws.Cells(lastRow, 2).Value = lstEntry.List(i, 1)
        ws.Cells(lastRow, 3).Value = lstEntry.List(i, 2)
        ws.Cells(lastRow, 4).Value = lstEntry.List(i, 3)
        ws.Cells(lastRow, 5).Value = lstEntry.List(i, 4)
        ws.Cells(lastRow, 6).Value = "OPEN"
        ws.Cells(lastRow, 7).Value = ""
        ws.Cells(lastRow, 8).Value = ""

        lastRow = lastRow + 1

    Next i

    Application.ScreenUpdating = True

    lstEntry.Clear

    MsgBox "Data NG berhasil disimpan.", vbInformation

End Sub
=================================================================================================================================================================================

Private Sub btnLoad_Click()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim resultArr() As Variant
    Dim i As Long, r As Long
    Dim monthNum As Integer

    Set ws = ThisWorkbook.Sheets("NG_Database")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    dataArr = ws.Range("A2:H" & lastRow).Value

    If cmbFilterMonth.ListIndex <> -1 Then
        monthNum = cmbFilterMonth.ListIndex + 1
    Else
        monthNum = 0
    End If

    ReDim resultArr(1 To UBound(dataArr), 1 To 6)
    r = 0

    For i = 1 To UBound(dataArr)

        If (cmbFilterSection.Value = "" Or dataArr(i, 2) = cmbFilterSection.Value) _
        And (cmbFilterStatus.Value = "ALL" Or dataArr(i, 6) = cmbFilterStatus.Value) _
        And (monthNum = 0 Or Month(dataArr(i, 1)) = monthNum) Then

            r = r + 1
            resultArr(r, 1) = i + 1
            resultArr(r, 2) = dataArr(i, 1)
            resultArr(r, 3) = dataArr(i, 2)
            resultArr(r, 4) = dataArr(i, 3)
            resultArr(r, 5) = dataArr(i, 5)
            resultArr(r, 6) = dataArr(i, 6)

        End If

    Next i

    If r > 0 Then
        ReDim Preserve resultArr(1 To r, 1 To 6)
        lstManage.List = resultArr
    Else
        lstManage.Clear
    End If

End Sub
=================================================================================================================================================================================

Private Sub lstManage_Click()

    If lstManage.ListIndex = -1 Then Exit Sub

    Dim ws As Worksheet
    Dim rowSheet As Long

    Set ws = ThisWorkbook.Sheets("NG_Database")

    rowSheet = lstManage.List(lstManage.ListIndex, 0)
    txtRowHidden.Value = rowSheet

    txtDateM.Value = ws.Cells(rowSheet, 1).Value
    txtSectionM.Value = ws.Cells(rowSheet, 2).Value
    txtParameterM.Value = ws.Cells(rowSheet, 3).Value
    txtDescM.Value = ws.Cells(rowSheet, 4).Value
    txtQtyM.Value = ws.Cells(rowSheet, 5).Value
    cmbStatusM.Value = ws.Cells(rowSheet, 6).Value
    txtActionM.Value = ws.Cells(rowSheet, 7).Value
    txtActionDateM.Value = ws.Cells(rowSheet, 8).Value

End Sub
=================================================================================================================================================================================

Private Sub btnUpdate_Click()

    Dim ws As Worksheet
    Dim rowSheet As Long

    If txtRowHidden.Value = "" Then
        MsgBox "Pilih data terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    If cmbStatusM.Value = "CLOSE" Then
        If Trim(txtActionM.Value) = "" Then
            MsgBox "Action wajib diisi sebelum closing.", vbExclamation
            txtActionM.SetFocus
            Exit Sub
        End If
    End If

    Set ws = ThisWorkbook.Sheets("NG_Database")

    rowSheet = CLng(txtRowHidden.Value)

    ws.Cells(rowSheet, 6).Value = cmbStatusM.Value
    ws.Cells(rowSheet, 7).Value = txtActionM.Value
    ws.Cells(rowSheet, 8).Value = txtActionDateM.Value

    MsgBox "Data berhasil diperbarui.", vbInformation

    Call btnLoad_Click
    Call ClearDetailFields

End Sub
=================================================================================================================================================================================

Private Sub btnClearDetail_Click()
    ClearDetailFields
End Sub


Private Sub ClearDetailFields()

    txtRowHidden.Value = ""
    txtDateM.Value = ""
    txtSectionM.Value = ""
    txtParameterM.Value = ""
    txtDescM.Value = ""
    txtQtyM.Value = ""
    cmbStatusM.Value = ""
    txtActionM.Value = ""
    txtActionDateM.Value = ""

End Sub
