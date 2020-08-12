VERSION 5.00
Begin VB.Form ATM 
   Caption         =   "Transaksi di ATM"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4305
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
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Selesai"
      Height          =   500
      Left            =   2160
      TabIndex        =   17
      Top             =   1800
      Width           =   2000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ganti PIN"
      Height          =   500
      Left            =   2160
      TabIndex        =   16
      Top             =   1320
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bayar Telepon"
      Height          =   500
      Left            =   2160
      TabIndex        =   15
      Top             =   840
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Transfer Antar Bank"
      Height          =   500
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Penarikan Tunai"
      Height          =   500
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Informasi Saldo"
      Height          =   500
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2000
   End
   Begin VB.TextBox TxtPin 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   3240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " Ketik Nomor PIN Anda ( 6 Digit )"
      Height          =   225
      Left            =   945
      TabIndex        =   11
      Top             =   120
      Width           =   2625
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No ATM"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Rekng"
      Height          =   345
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Kartu"
      Height          =   345
      Left            =   2040
      TabIndex        =   6
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Label LblNoATM 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Label LblJam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label LblNoRek 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label LblNoKartu 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1005
   End
End
Attribute VB_Name = "ATM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call BukaDB
RSNasabah.Open "select * from nasabah where norek='" & LblNoRek & "'", Conn
If Not RSNasabah.EOF Then
    MsgBox "Nama  : " & RSNasabah!namansb & "" & Chr(13) & _
    "Saldo  : Rp. " & Format(RSNasabah!Saldo, "###,###,###") & ""
End If
End Sub

Private Sub Command2_Click()
Ambil1.Show
End Sub

Private Sub Command3_Click()
Transfer.Show
End Sub

Private Sub Command4_Click()
Pembayaran.Show
End Sub

Private Sub Command5_Click()
GantiPIN.Show
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
TxtPin.PasswordChar = "X"
TxtPin = ""
TxtPin.MaxLength = 6
LblTanggal = Date
LblNoATM = "ATM01"
End Sub

Private Sub Timer1_Timer()
LblJam = Time$
End Sub

Private Sub TxtPin_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Unload Me
End If

If Keyascii = 13 Then
    If Len(TxtPin) <> 6 Then
        MsgBox "Nomor Pin Harus 6 digit"
        TxtPin = ""
        TxtPin.SetFocus
        Exit Sub
    End If
    Call BukaDB
    RSNasabah.Open "Select * from Nasabah where PIN='" & TxtPin & "'", Conn
    If Not RSNasabah.EOF Then
        TxtPin.Enabled = False
        Command1.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        LblNoRek = RSNasabah!norek
        Transfer.Pengirim = RSNasabah!namansb
    Else
        Pesan = MsgBox("Nomor PIN << " & TxtPin & " >> Salah", 0, "Peringatan")
        TxtPin = ""
        TxtPin.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub lblnorek_Change()
Call BukaDB
RSNasabah.Open "select * from nasabah where norek='" & LblNoRek & "'", Conn
If Not RSNasabah.EOF Then
    LblNoKartu = RSNasabah!Nokartu
    AmbilAtm.LblNama = RSNasabah!namansb
End If
End Sub

