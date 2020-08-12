VERSION 5.00
Begin VB.Form Pengambilan 
   Caption         =   "Pengambilan Kas"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Width           =   800
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   840
   End
   Begin VB.TextBox TxtJumlahAbl 
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   1500
   End
   Begin VB.TextBox TxtNoRek 
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   3000
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Rekening"
      Height          =   345
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nasabah"
      Height          =   345
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Penarikan"
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label LblJam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label LblKodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3840
      TabIndex        =   15
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   14
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   13
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label LblNamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3840
      TabIndex        =   12
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Kasir"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2760
      TabIndex        =   11
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Transaksi"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      Height          =   345
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label LblSaldo 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label LblNama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "Pengambilan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LblKodeKsr = Login.TxtKodeKsr
LblNamaKsr = Login.TxtNamaKsr
LblTanggal = Date
TxtJumlahAbl.MaxLength = 8
TidakSiapIsi
End Sub

Private Sub Form_Activate()
If LblKodeKsr = "" Then
    MsgBox "Kode kasir tidak terdeteksi"
    Login.Show
End If

Call BukaDB
RSNasabah.Open "select * from nasabah", Conn
RSNasabah.Requery
Combo1.Clear
Do While Not RSNasabah.EOF
    Combo1.AddItem RSNasabah!norek
    RSNasabah.MoveNext
Loop

Call Auto
End Sub

Private Sub Timer1_Timer()
LblJam = Time$
End Sub

Function CariNasabah()
Call BukaDB
RSNasabah.Open "Select * From Nasabah where NoRek='" & TxtNoRek & "'", Conn
End Function

Function CariCombo()
Call BukaDB
RSNasabah.Open "Select * From Nasabah where NoRek='" & Combo1 & "'", Conn
End Function

Function CariKasir()
Call BukaDB
RSKasir.Open "Select * From Kasir where KodeKsr='" & LblKodeKsr & "'"
End Function

Private Sub CmdInput_Click()
If CmdInput.Caption = "&Input" Then
    CmdInput.Caption = "&Simpan"
    CmdTutup.Caption = "&Batal"
    SiapIsi
    Combo1.SetFocus
Else
    If TxtNoRek = "" Or TxtJumlahAbl = "" Then
        MsgBox "Data Belum lengkap", 0, "Periksa Kembali Isian Data"
    Else
    
        Dim SimpanTransaksi As String
        SimpanTransaksi = "Insert into Transaksi(NoTransaksi,Norek,Tanggal,Jam,Pengeluaran,Keterangan,KodeKsr) values " & _
        "('" & Nomor & "','" & TxtNoRek & "','" & LblTanggal & "','" & LblJam & "','" & TxtJumlahAbl & "','" & Pengambilan.Caption & "','" & Menu.StatusBar1.Panels(1).Text & "')"
        Conn.Execute SimpanTransaksi
                
        Call CariNasabah
        If Not RSNasabah.EOF Then
            Dim Edit As String
            Edit = "Update nasabah set saldo='" & RSNasabah!Saldo - Val(TxtJumlahAbl) & "' where Norek='" & TxtNoRek & "'"
            Conn.Execute Edit
            Kosongkan
            Semula
            TidakSiapIsi
        End If
        Form_Activate
        Menu.DT.Refresh
        CetakLayar
    End If
End If
End Sub

Private Sub TxtNoRek_KeyPress(Keyascii As Integer)
On Error Resume Next
If Keyascii = 27 Then
    Semula
    CmdTutup.SetFocus
End If
If Keyascii = 13 Then
    If CmdInput.Caption = "&Simpan" Then
        Call CariNasabah
        If Not RSNasabah.EOF Then
            LblNama = RSNasabah!namansb
            LblSaldo = RSNasabah!Saldo
            LblSaldo = Format(RSNasabah!Saldo, "###,###,###")
            TxtJumlahAbl.SetFocus
        Else
            X = MsgBox("Rekening No : < " & TxtNoRek & " > tidak ada, coba nomor lain...!", "Informasi")
            TxtNoRek.SetFocus
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TxtJumlahAbl_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Semula
    CmdTutup.SetFocus
End If
If Keyascii = 13 Then
    If Val(TxtJumlahAbl) > RSNasabah!Saldo Then
        MsgBox "Dana tidak cukup" & Chr(13) & _
        "Saldo hanya Rp. " & Format(RSNasabah!Saldo, "###,###,###") & ""
        TxtJumlahAbl = ""
        TxtJumlahAbl.SetFocus
        Exit Sub
    End If

    If CmdInput.Caption = "&Simpan" Then
        CmdInput.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub CmdTutup_Click()
Select Case CmdTutup.Caption
    Case "&Tutup"
        Unload Me
    Case "&Batal"
        TidakSiapIsi
        Semula
        Kosongkan
End Select
End Sub

Sub TidakSiapIsi()
TxtNoRek.Enabled = False
TxtJumlahAbl.Enabled = False
End Sub

Sub SiapIsi()
TxtNoRek.Enabled = True
TxtJumlahAbl.Enabled = True
End Sub

Sub Kosongkan()
TxtNoRek = ""
LblNama = ""
LblSaldo = ""
TxtJumlahAbl = ""
End Sub

Sub Semula()
Kosongkan
TidakSiapIsi
CmdInput.Caption = "&Input"
CmdTutup.Caption = "&Tutup"
End Sub


Sub CetakLayar()
Tampilkan.Show
Tampilkan.Font = "Courier New"
Call BukaDB
RSTransaksi.Open "select * from Transaksi Where NoTransaksi In(Select Max(NoTransaksi)From Transaksi)Order By NoTransaksi Desc", Conn
RSNasabah.Open "select * from Nasabah where NoRek='" & RSTransaksi!norek & "'", Conn
RSKasir.Open "select * from Kasir where KodeKsr='" & RSTransaksi!kodeksr & "'", Conn
Tampilkan.Print
Tampilkan.Print Tab(5); "Bukti Penarikan (Teller)"
Tampilkan.Print
Tampilkan.Print Tab(5); "Nomor    :  "; RSTransaksi!NoTransaksi
Tampilkan.Print Tab(5); "Tanggal  :  "; RSTransaksi!tanggal
Tampilkan.Print Tab(5); "Jam      :  "; RSTransaksi!Jam
Tampilkan.Print Tab(5); "Kasir    :  "; RSKasir!NamaKsr
Tampilkan.Print Tab(5); "Nasabah  :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "Jumlah   :  "; RKanan(RSTransaksi!Pengeluaran, "###,###,###")
Tampilkan.Print Tab(5); "Saldo    :  "; RKanan(RSNasabah!Saldo, "###,###,###")
End Sub

Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Private Sub Combo1_Click()
Call BukaDB
RSNasabah.Open "select * from nasabah where norek='" & Combo1 & "'", Conn
If Not RSNasabah.EOF Then
    With RSNasabah
    If Not RSNasabah.EOF Then
        TxtNoRek = RSNasabah!norek
        LblNama = RSNasabah!namansb
        LblSaldo = RSNasabah!Saldo
    End If
    End With
End If
End Sub

Private Sub Combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If CmdInput.Caption = "&Simpan" Then
        Call CariCombo
        If Not RSNasabah.EOF Then
            TxtNoRek.Enabled = False
            TxtJumlahAbl.SetFocus
        Else
            MsgBox "Nomor Rekening Nasabah Tidak Ada"
            TxtNoRek.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

