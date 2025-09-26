Option Explicit

Private arrSizes As Variant   ' cache kolom A (A2..lastRow)

'========== INIT =========='
Private Sub UserForm_Initialize()
    Dim ws As Worksheet, lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("data_spec")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    arrSizes = ws.Range("A2:A" & lastRow).Value2
    If Not IsArray(arrSizes) Then
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = ws.Range("A2").Value2
        arrSizes = tmp
    End If
    
    ' awal: tampilkan semua
    RenderByDigits vbNullString
End Sub

'========== FILTER RENDER =========='
' digits = "250", "249", "" (kosong = tampil semua)
Private Sub RenderByDigits(ByVal digits As String)
    Dim i As Long, n As Long, s As String
    Dim target As String, needLen As Long
    
    ComboBoxSizeGT.Clear
    If IsEmpty(arrSizes) Then Exit Sub
    
    ' Normalisasi input: "250" -> target "G250"
    digits = Trim$(digits)
    ' kalau kosong → tampil semua
    If Len(digits) = 0 Then
        For i = 1 To UBound(arrSizes, 1)
            s = CStr(arrSizes(i, 1))
            If Len(s) > 0 Then ComboBoxSizeGT.AddItem s
        Next
        Exit Sub
    End If
    
    ' Boleh ketik "G250" atau "250". Samakan jadi "G" + angka.
    If UCase$(Left$(digits, 1)) = "G" Then
        digits = Mid$(digits, 2)  ' buang huruf G di depan jika ada
    End If
    
    ' target prefix bertahap:
    ' - ketik 1 digit: "2"  -> cocokkan Left(s,2)  = "G2"
    ' - ketik 2 digit: "25" -> cocokkan Left(s,3)  = "G25"
    ' - ketik 3 digit: "250"-> cocokkan Left(s,4)  = "G250"
    If Len(digits) > 3 Then digits = Left$(digits, 3)
    
    target = "G" & digits
    needLen = 1 + Len(digits)           ' panjang prefix yang harus sama
    
    For i = 1 To UBound(arrSizes, 1)
        s = CStr(arrSizes(i, 1))
        If Len(s) >= needLen Then
            If Left$(s, needLen) = target Then
                ComboBoxSizeGT.AddItem s
            End If
        End If
    Next i
    
    ' Jika kosong karena user ketik angka aneh → tidak isi apa-apa.
End Sub

'========== EVENT: user mengetik di TextBox =========='
Private Sub TextBoxFilterSize_Change()
    ' bersihkan semua non-digit dan non-G (kalau user ngetik "G250")
    Dim t As String, i As Long, ch As String
    t = TextBoxFilterSize.Text
    Dim cleaned As String
    
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "#" Or UCase$(ch) = "G" Then cleaned = cleaned & ch
    Next i
    
    If cleaned <> t Then
        ' set balik tanpa pindahkan caret ke depan
        Dim pos As Long: pos = TextBoxFilterSize.SelStart
        TextBoxFilterSize.Text = cleaned
        TextBoxFilterSize.SelStart = pos
    End If
    
    ' ambil hanya angka untuk logika filter ("250" dari "G250")
    Dim digits As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "#" Then digits = digits & ch
    Next i
    
    ' render hasil
    RenderByDigits digits
End Sub

'========== OPTIONAL: batasi input ke angka & G saat KeyPress =========='
Private Sub TextBoxFilterSize_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim c As Integer: c = KeyAscii
    If c = vbKeyBack Then Exit Sub
    If (c >= 48 And c <= 57) Or UCase$(Chr$(c)) = "G" Then
        ' boleh angka 0-9 dan huruf G
    Else
        KeyAscii = 0
    End If
End Sub
