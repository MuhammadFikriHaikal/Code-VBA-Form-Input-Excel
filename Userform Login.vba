Private Sub CheckBox1_Click()
    If CheckBox1 = True Then
        TextBox1.PasswordChar = ""
    Else: TextBox1.PasswordChar = "*"
    End If
End Sub

Private Sub CheckBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckBox1.ForeColor = &H8000000F
End Sub

Private Sub CommandButton1_Click()
    CommandButton1_Enter
End Sub

Private Sub CommandButton1_Enter()
    Dim NamaPengguna As String, Sandi As String
        Dim Namasaya1 As String, Sandisaya1 As String
        Dim Namasaya2 As String, Sandisaya2 As String
        Dim Namasaya3 As String, Sandisaya3 As String
    
        NamaPengguna = TextBoxNama.Text
        Sandi = TextBox1.Text
        Namasaya1 = "fikri"
        Namasaya2 = "nabila"
        Namasaya3 = "yehes"
        Sandisaya1 = "fikri"
        Sandisaya2 = "nabila"
        Sandisaya3 = "yehes"
    
        If NamaPengguna = Empty Then
            TextBoxNama.SetFocus
        ElseIf Sandi = Empty Then
            MsgBox "Silahkan Isi Password", vbExclamation + vbOKOnly, "Pemberitahuan"
            TextBox1.SetFocus
        ElseIf NamaPengguna = Namasaya1 And Sandi = Sandisaya1 Then
            MsgBox "Login berhasil..!", vbInformation + vbOKOnly, "Pemberitahuan"
            Sheets("Form Input Kas").Select
            Unload Me
            Call Users1
        ElseIf NamaPengguna = Namasaya2 And Sandi = Sandisaya2 Then
            MsgBox "Login berhasil..!", vbInformation + vbOKOnly, "Pemberitahuan"
            Sheets("Form Input Kas").Select
            Unload Me
            Call Users2
        ElseIf NamaPengguna = Namasaya3 And Sandi = Sandisaya3 Then
            MsgBox "Login berhasil..!", vbInformation + vbOKOnly, "Pemberitahuan"
            Sheets("Form Input Kas").Select
            Unload Me
            Call Users3
        Else
            MsgBox "Nama Pengguna atau Sandi Salah!", vbCritical + vbOKOnly, "Peringatan"
            TextBoxNama.Text = ""
            TextBox1.Text = ""
            TextBoxNama.SetFocus
        End If
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton1.BackColor = &HFF00&
    TombolMas.BackColor = &H8000000F
End Sub

Private Sub Label12_Click()
    Dim NamaPengguna As String, Sandi As String
    
    NamaPengguna = TextBoxNama.Text
    Sandi = TextBox1.Text
    
    If NamaPengguna = Empty Then
        MsgBox "Silahkan Isi Username Terlebih Dahulu..", vbInformation, "Dari Excel"
    ElseIf NamaPengguna = "fikri" Then
        TextBox1.Text = "fikri"
        MsgBox "Dah Tuhh..", vbInformation, "Bantuan Excel"
    ElseIf NamaPengguna = "nabila" Then
        TextBox1.Text = "nabila"
        MsgBox "Dah Tuhh..", vbInformation, "Bantuan Excel"
    ElseIf NamaPengguna = "yehes" Then
        TextBox1.Text = "yehes"
        MsgBox "Dah Tuhh..", vbInformation, "Bantuan Excel"
    Else
        MsgBox "Username Tidak Dikenal..", vbCritical, "dari excel"
        TextBoxNama.SetFocus
    End If
End Sub

Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label12.ForeColor = &H8000000F
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton1.BackColor = &H8000000F
    TombolMas.BackColor = &H8000000F
    CheckBox1.ForeColor = &H80000006
    Label12.ForeColor = &H80000006
End Sub

Private Sub TextBoxNama_Enter()
    TextBoxNama.BorderColor = &HFF00&
End Sub

Private Sub TextBoxNama_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBoxNama.BorderColor = &H80000006
End Sub

Private Sub TombolMas_Click()
    strAnswer = MsgBox("Yakin gajadi login..?", vbQuestion + vbOKCancel, "Dari Excel")
    If strAnswer = vbOK Then
        Application.ThisWorkbook.Close
    End If
End Sub

Private Sub TombolMas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TombolMas.BackColor = &H80000006
    CommandButton1.BackColor = &H8000000F
End Sub

Private Sub UserForm_Activate()
    With Application
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
        TextBoxNama.SetFocus
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Application
        Me.Top = .Top
        Me.Left = .Left
        Me.Height = .Height
        Me.Width = .Width
    End With
End Sub

Private Sub UserForm_Initialize()
    'SetMaxMin (Me.Caption)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        MsgBox "Silahkan masukan Username dan Password untuk login, atau tekan Cancel untuk keluar", vbInformation, "Dari Excel"
    End If
End Sub

' LOGIC LOGIN BARU ADMIN/TAMU
Private Sub cmdLogin_Click()
    Dim username As String, password As String
    Dim userFile As String
    Dim ws As Worksheet
    Dim foundRow As Range
    Dim userRole As String

    ' File dan sheet data pengguna
    userFile = "C:\Users\Username\Documents\UserData.xlsx" ' Ganti dengan lokasi file Anda

    ' Ambil input dari TextBox
    username = TextBoxUsername.Value
    password = TextBoxPassword.Value

    ' Cek apakah file ada
    If Dir(userFile) = "" Then
        MsgBox "File data pengguna tidak ditemukan!", vbCritical
        Exit Sub
    End If

    ' Buka file pengguna
    Dim wb As Workbook
    Set wb = Workbooks.Open(userFile)
    Set ws = wb.Sheets(1) ' Asumsi data ada di sheet pertama

    ' Cari username di file pengguna
    Set foundRow = ws.Columns("A").Find(What:=username, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundRow Is Nothing Then
        ' Cek password
        If foundRow.Offset(0, 1).Value = password Then
            ' Ambil peran pengguna
            userRole = foundRow.Offset(0, 2).Value

            ' Tutup file pengguna
            wb.Close SaveChanges:=False

            ' Buka dashboard sesuai peran
            UserFormDashboard.InitializeAccess userRole
            UserFormDashboard.Show
            Unload Me
        Else
            MsgBox "Password salah!", vbExclamation
            wb.Close SaveChanges:=False
        End If
    Else
        MsgBox "Username tidak ditemukan!", vbExclamation
        wb.Close SaveChanges:=False
    End If
End Sub

' LOGIC USERNAME & PASSWORD MENGGUNAKAN SELECT CASE
Function GetLogin(User As String, Pass As String) As Boolean
    Dim Passsword As String
        Select Case User
            Case "fikri": Password = "fikri"
            Case "nabila": Password = "nabila"
            Case "yehes": Password = "yehes"
            Case Else: GetLogin = False
        End Select
            GetLogin = IIf(Password = Pass, True, False)
End Function

' CODE UNTUK DI USERFORM DASHBOARDNYA
Public Sub InitializeAccess(ByVal role As String)
    If role = "Admin" Then
        EnableEditMode True
        MsgBox "Selamat datang, Admin!", vbInformation
    ElseIf role = "Tamu" Then
        EnableEditMode False
        MsgBox "Selamat datang, Tamu! Anda hanya dapat membaca data.", vbInformation
    End If
End Sub

Private Sub EnableEditMode(ByVal canEdit As Boolean)
    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = Not canEdit
        ElseIf TypeOf ctrl Is CommandButton Then
            If ctrl.Name = "CommandButtonAdd" Then
                ctrl.Enabled = canEdit
            End If
        End If
    Next ctrl
End Sub
