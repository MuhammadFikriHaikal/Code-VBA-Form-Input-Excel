Option Explicit

' Cache data agar render cepat dan konsisten 2D meski 1 baris
Private arrSizes As Variant   ' 2D array kolom A (A2..lastRow)

'================ UTIL: cari last row aman ================'
Private Function GetLastRowSafe(ws As Worksheet, Optional ByVal col As String = "A") As Long
    Dim r As Range
    On Error Resume Next
    Set r = ws.Columns(col).Find(What:="*", LookIn:=xlFormulas, _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If r Is Nothing Then
        GetLastRowSafe = 1
    Else
        GetLastRowSafe = r.Row
    End If
End Function

'================ QuickSort string sederhana ================'
Private Sub QuickSortStrings(arr As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, t As String
    i = lo: j = hi: p = arr((lo + hi) \ 2)
    Do While i <= j
        Do While arr(i) < p: i = i + 1: Loop
        Do While arr(j) > p: j = j - 1: Loop
        If i <= j Then t = arr(i): arr(i) = arr(j): arr(j) = t: i = i + 1: j = j - 1
    Loop
    If lo < j Then QuickSortStrings arr, lo, j
    If i < hi Then QuickSortStrings arr, i, hi
End Sub

'================ INIT ================='
Private Sub UserForm_Initialize()
    Dim ws As Worksheet, lastRow As Long
    Dim dictPrefix As Object
    Dim i As Long, s As String, p As String
    Dim keys As Variant

    ' GANTI jika nama sheet beda
    Set ws = ThisWorkbook.Worksheets("data_spec")

    lastRow = GetLastRowSafe(ws, "A")
    If lastRow < 2 Then
        ComboBoxSizeGT.Clear
        ComboBoxFilterSize.Clear
        Exit Sub
    End If

    ' Muat data kolom A
    arrSizes = ws.Range("A2:A" & lastRow).Value2
    ' Normalisasi jika cuma 1 baris
    If Not IsArray(arrSizes) Then
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = ws.Range("A2").Value2
        arrSizes = tmp
    End If

    ' Kumpulkan prefix 4-karakter unik
    Set dictPrefix = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arrSizes, 1)
        s = CStr(arrSizes(i, 1))
        If Len(s) >= 4 Then
            p = Left$(s, 4)            ' contoh: G250JT-01 -> G250
            If Not dictPrefix.Exists(p) Then dictPrefix.Add p, 1
        End If
    Next i

    ' Isi ComboBoxFilterSize (urut)
    keys = dictPrefix.Keys
    If IsArray(keys) Then
        QuickSortStrings keys, LBound(keys), UBound(keys)
        ComboBoxFilterSize.Clear
        ComboBoxFilterSize.List = keys
    End If

    ' Tampilkan semua dulu
    RenderSizesByPrefix vbNullString
End Sub

'================ RENDER LIST ================='
Private Sub RenderSizesByPrefix(ByVal prefix As String)
    Dim i As Long, s As String
    Dim useFilter As Boolean

    ComboBoxSizeGT.Clear
    If IsEmpty(arrSizes) Then Exit Sub

    prefix = Trim$(prefix)
    useFilter = (Len(prefix) > 0)

    For i = 1 To UBound(arrSizes, 1)
        s = CStr(arrSizes(i, 1))
        If s <> vbNullString Then
            If useFilter Then
                If Len(s) >= 4 Then
                    If Left$(s, 4) = prefix Then ComboBoxSizeGT.AddItem s
                End If
            Else
                ComboBoxSizeGT.AddItem s
            End If
        End If
    Next i
End Sub

'================ EVENT FILTER ================='
Private Sub ComboBoxFilterSize_Change()
    RenderSizesByPrefix ComboBoxFilterSize.Text
End Sub

Private Sub ComboBoxFilterSize_Click()
    RenderSizesByPrefix ComboBoxFilterSize.Text
End Sub
