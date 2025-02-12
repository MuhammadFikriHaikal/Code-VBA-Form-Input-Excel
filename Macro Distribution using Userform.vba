Private Sub CmdSave_Click()
    
    ' Deklarasi Workbook dan Worksheet
    Dim wbDatabase As Workbook
    Dim wsDatabase As Worksheet
    Dim targetPath As String
    Dim dbLastRow As Long
    Dim checkBoxes As Variant
    Dim numFilledChecks As Integer
    Dim i As Integer
    Dim key As Variant

    ' Path file database
    targetPath = "C:\Users\michael\Documents\daily_database.xlsx"

    ' Buka file daily_database jika belum terbuka
    Set wbDatabase = GetWorkbook(targetPath)
    If wbDatabase Is Nothing Then Exit Sub ' Jika gagal membuka, hentikan

    ' Set worksheet
    Set wsDatabase = wbDatabase.Sheets("Daily Database")

    ' TextBox pengecekan (misalnya untuk pengecekan mold atau suhu)
    checkBoxes = Array("TextBoxP1", "TextBoxP2", "TextBoxP3") ' Ganti dengan nama TextBox Anda

    ' Hitung jumlah TextBox pengecekan yang terisi
    numFilledChecks = 0
    For i = LBound(checkBoxes) To UBound(checkBoxes)
        If Trim(Me.Controls(checkBoxes(i)).Value) <> "" Then
            numFilledChecks = numFilledChecks + 1
        End If
    Next i

    ' Jika tidak ada pengecekan yang terisi, beri pesan dan hentikan
    If numFilledChecks = 0 Then
        MsgBox "Tidak ada hasil pengecekan yang terisi. Data tidak akan disimpan.", vbExclamation
        Exit Sub
    End If

    ' TextBox data utama (Date, Group, Shift, Machine, Size)
    Dim mainDataBoxes As Variant
    mainDataBoxes = Array("TextBoxDate", "TextBoxGroup", "TextBoxShift", "TextBoxMachine")

    ' Validasi input untuk TextBox data utama
    For i = LBound(mainDataBoxes) To UBound(mainDataBoxes)
        If Trim(Me.Controls(mainDataBoxes(i)).Value) = "" Then
            MsgBox "Harap isi semua data utama sebelum menyimpan!", vbExclamation
            Exit Sub
        End If
    Next i

    ' Salin data utama ke worksheet sebanyak jumlah pengecekan yang terisi
    With wsDatabase
        dbLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1 ' Cari baris terakhir di kolom A

        For i = 1 To numFilledChecks
            .Cells(dbLastRow, 1).Value = Me.TextBoxDate.Value ' Kolom Date
            .Cells(dbLastRow, 2).Value = Me.TextBoxGroup.Value ' Kolom Group
            .Cells(dbLastRow, 3).Value = Me.TextBoxShift.Value ' Kolom Shift
            .Cells(dbLastRow, 4).Value = Me.TextBoxMachine.Value ' Kolom Machine
            .Cells(dbLastRow, 5).Value = Me.ComboBoxSizeGT.Value
            dbLastRow = dbLastRow + 1 ' Pindah ke baris berikutnya
        Next i
    End With

    ' Mapping TextBox ke kolom target (untuk data pengecekan lainnya)
    Dim mapping As Object
    Set mapping = CreateObject("Scripting.Dictionary")
    mapping.Add "TextBoxP1", "F" ' TextBoxP1 -> Kolom F
    mapping.Add "TextBoxP2", "F" ' TextBoxP2 -> Kolom G
    mapping.Add "TextBoxP3", "F" ' TextBoxP3 -> Kolom H
    ' Tambahkan lebih banyak TextBox sesuai kebutuhan...

    ' Pindahkan data pengecekan ke worksheet
    With wsDatabase
        For Each key In mapping.Keys
            dbLastRow = .Cells(.Rows.Count, mapping(key)).End(xlUp).Row + 1
            .Cells(dbLastRow, mapping(key)).Value = Me.Controls(key).Value
        Next key
    End With

    ' Simpan workbook
    wbDatabase.Save

    ' Bersihkan TextBox setelah save
    Dim excludeList As Object
    Set excludeList = CreateObject("Scripting.Dictionary")
    excludeList.Add "TextBoxGroup", True ' Contoh: TextBoxGroup tidak akan dihapus
    excludeList.Add "TextBoxShift", True ' Contoh: TextBoxShift tidak akan dihapus

    For Each key In mainDataBoxes
        If Not excludeList.Exists(key) Then
            Me.Controls(key).Value = ""
        End If
    Next key

    For Each key In checkBoxes
        Me.Controls(key).Value = ""
    Next key

    MsgBox "Data berhasil disimpan sebanyak " & numFilledChecks & " kali!", vbInformation
    Exit Sub ' Keluar dari prosedur
End Sub

' Fungsi untuk membuka workbook
Private Function GetWorkbook(targetPath As String) As Workbook
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks.Open(targetPath)
    On Error GoTo 0

    If wb Is Nothing Then
        MsgBox "Gagal membuka file " & targetPath, vbCritical
    End If
    Set GetWorkbook = wb
End Function

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    ' Referensi ke worksheet dan range data
    Set ws = ThisWorkbook.Sheets("data_spec")
    Set rng = ws.Range("A2:A5")

    ' Tambahkan data dari range ke ComboBox
    For Each cell In rng
        Me.ComboBoxSizeGT.AddItem cell.Value
    Next cell
End Sub

Private Sub ComboBoxSizeGT_Change()
    Dim ws As Worksheet
    Dim rng As Range
    Dim sizeRow As Range
    Dim selectedSize As String
    Dim mapping As Object
    Dim key As Variant

    ' Referensi worksheet sumber data
    Set ws = ThisWorkbook.Sheets("data_spec")
    Set rng = ws.Range("A2:L5") ' Pastikan range mencakup semua data Anda

    ' Ambil nilai dari ComboBox
    selectedSize = Me.ComboBoxSizeGT.Value

    ' Buat mapping antara Label dan kolom
    Set mapping = CreateObject("Scripting.Dictionary")
    mapping.Add "Label1", 2  ' Kolom B
    mapping.Add "Label2", 3  ' Kolom D
    mapping.Add "Label3", 4  ' Kolom F
    ' Tambahkan mapping lainnya sesuai kebutuhan...

    ' Cari baris data untuk Size yang dipilih
    For Each sizeRow In rng.Columns(1).Cells
        If sizeRow.Value = selectedSize Then
            ' Tampilkan data di Label sesuai mapping
            For Each key In mapping.Keys
                Me.Controls(key).Caption = sizeRow.Offset(0, mapping(key) - 1).Value
            Next key
            Exit Sub
        End If
    Next sizeRow

    ' Jika Size tidak ditemukan, kosongkan label
    For Each key In mapping.Keys
        Me.Controls(key).Caption = "Data tidak ditemukan"
    Next key
End Sub
