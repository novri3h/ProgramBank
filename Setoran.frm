VERSION 5.00
Begin VB.Form Setoran 
   Caption         =   "Setoran Kas"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
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
   ScaleHeight     =   3570
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   3000
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1860
   End
   Begin VB.TextBox TxtNoRek 
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin VB.TextBox TxtJumlahStr 
      Height          =   350
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   840
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   800
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   800
   End
   Begin VB.Label LblNama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   19
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label LblSaldo 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   18
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      Height          =   345
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Transaksi"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1245
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
      TabIndex        =   13
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label LblNamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3840
      TabIndex        =   12
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   10
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label LblKodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label LblJam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   8
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Setoran"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1250
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1250
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nasabah"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1250
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Rekening"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1250
   End
End
Attribute VB_Name = "Setoran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LblKodeKsr = Login.TxtKodeKsr
LblNamaKsr = Login.TxtNamaKsr
LblTanggal = Date
TxtJumlahStr.MaxLength = 8
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
    'TxtNoRek.SetFocus
    Combo1.SetFocus
Else
    If TxtNoRek = "" Or TxtJumlahStr = "" Then
        MsgBox "Data Belum lengkap", 0, "Periksa Kembali Isian Data"
    Else
    
        Dim SimpanTransaksi As String
        SimpanTransaksi = "Insert into Transaksi(NoTransaksi,Norek,Tanggal,Jam,Pemasukan,Keterangan,KodeKsr) values " & _
        "('" & Nomor & "','" & TxtNoRek & "','" & LblTanggal & "','" & LblJam & "','" & TxtJumlahStr & "','" & Setoran.Caption & "','" & Menu.StatusBar1.Panels(1).Text & "')"
        Conn.Execute SimpanTransaksi
                
        Call CariNasabah
        If Not RSNasabah.EOF Then
            Dim Edit As String
            Edit = "Update nasabah set saldo='" & RSNasabah!Saldo + Val(TxtJumlahStr) & "' where Norek='" & TxtNoRek & "'"
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
            TxtJumlahStr.SetFocus
        Else
            X = MsgBox("Rekening No : < " & TxtNoRek & " > tidak ada, coba nomor lain...!", "Informasi")
            TxtNoRek.SetFocus
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TxtJumlahStr_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Semula
    CmdTutup.SetFocus
End If
If Keyascii = 13 Then
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
TxtJumlahStr.Enabled = False
End Sub

Sub SiapIsi()
TxtNoRek.Enabled = True
TxtJumlahStr.Enabled = True
End Sub

Sub Kosongkan()
TxtNoRek = ""
LblNama = ""
LblSaldo = ""
TxtJumlahStr = ""
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
Tampilkan.Print Tab(5); "Bukti Setoran (Teller)"
Tampilkan.Print
Tampilkan.Print Tab(5); "Nomor    :  "; RSTransaksi!NoTransaksi
Tampilkan.Print Tab(5); "Tanggal  :  "; RSTransaksi!tanggal
Tampilkan.Print Tab(5); "Jam      :  "; RSTransaksi!Jam
Tampilkan.Print Tab(5); "Kasir    :  "; RSKasir!NamaKsr
Tampilkan.Print Tab(5); "Nasabah  :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "Jumlah   :  "; RKanan(RSTransaksi!Pemasukan, "###,###,###")
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
            TxtJumlahStr.SetFocus
        Else
            MsgBox "Nomor Rekening Nasabah Tidak Ada"
            TxtNoRek.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

