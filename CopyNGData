Sub CopyNGData()
    ' deklarasi awal
    Dim wsDatabase As Worksheet
    Dim wsIS As Worksheet
    Dim wbDatabase As Workbook
    Dim wbIS As Workbook
    Dim sourcePathDatabase As String
    Dim targetPathIS As String
    Dim lastRowDatabase As Long
    Dim lastRowIS As Long
    Dim i As Long, j As Long

    ' Inisialisasi path file
    sourcePathDatabase = "C:\Users\michael\Documents\daily_database.xlsx"
    sourcePathIS = "C:\Users\michael\Documents\DB.xlsx"

    ' Inisialisasi worksheet
    Set wbDatabase = Workbooks.Open(sourcePathDatabase)
    Set wbIS = Workbooks.Open(sourcePathIS)
    
    Set wsDatabase = wbDatabase.Sheets("Daily Database")
    Set wsIS = wbIS.Sheets("Information Sheet")

    ' Cari baris terakhir di masing-masing sheet
    lastRowDatabase = wsDatabase.Cells(wsDatabase.Rows.Count, 1).End(xlUp).Row ' Kolom (date)
    lastRowIS = wsIS.Cells(wsIS.Rows.Count, 1).End(xlUp).Row          ' Kolom A di Sheet DB

    For i = 4 To lastRowDatabase ' Mulai dari baris 4
        For j = 5 To wsDatabase.Cells(2, wsDatabase.Columns.Count).End(xlToLeft).Column ' Baris 2 adalah jenis data
            ' Periksa apakah hasil NG
            If wsDatabase.Cells(i, j).Offset(0, 1).Value = "NG" Then
                ' Tambahkan baris baru di Sheet DB
                lastRowIS = lastRowIS + 1

                ' Copy data tambahan ke Sheet DB
                wsIS.Cells(lastRowIS, 1).Value = wsDatabase.Cells(i, 1).Value ' Date
                wsIS.Cells(lastRowIS, 2).Value = wsDatabase.Cells(i, 2).Value ' Shift
                wsIS.Cells(lastRowIS, 3).Value = wsDatabase.Cells(i, 3).Value ' Group
                wsIS.Cells(lastRowIS, 4).Value = wsDatabase.Cells(i, 4).Value ' Machine
                wsIS.Cells(lastRowIS, 5).Value = wsDatabase.Cells(i, 5).Value ' Size

                ' Tambahkan jenis data ke kolom "Jenis Data NG"
                wsIS.Cells(lastRowIS, 6).Value = wsDatabase.Cells(2, j).Value ' Jenis data

                ' Tambahkan subkategori (Left/Right)
                wsIS.Cells(lastRowIS, 7).Value = wsDatabase.Cells(3, j).Value ' Side nya

                ' Tambahkan nilai data NG
                wsIS.Cells(lastRowIS, 11).Value = wsDatabase.Cells(i, j).Value ' Data NG
            End If
        Next j
    Next i

    MsgBox "Data NG berhasil dicopy ke Sheet DB!", vbInformation
End Sub
