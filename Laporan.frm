VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Laporan 
   Caption         =   "Laporan Transaksi Per Kasir"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Laporan Pengambilan Hari Ini"
      Height          =   600
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Laporan Setoran Hari Ini"
      Height          =   600
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Laporan Nasabah Baru"
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3500
   End
   Begin VB.Label LblKodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   8
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label LblNamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   6
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Kasir"
      Height          =   345
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      Height          =   350
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1500
   End
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LblKodeKsr = Login.TxtKodeKsr
LblNamaKsr = Login.TxtNamaKsr
LblTanggal = Date
End Sub

Private Sub Command1_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command1_Click()
    CrystalReport1.SelectionFormula = "ToText({Transaksi.Tanggal})='" & Date & "' and {Transaksi.KodeKsr}='" & LblKodeKsr & "' and {Transaksi.Keterangan}='Buka Rekening'"
    CrystalReport1.ReportFileName = App.Path & "\Lap Nasabah Baru1.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub Command2_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command2_Click()
    CrystalReport1.SelectionFormula = "ToText({Transaksi.Tanggal})='" & Date & "' and {Transaksi.KodeKsr}='" & LblKodeKsr & "'and {Transaksi.Keterangan}='Setoran Kas'"
    CrystalReport1.ReportFileName = App.Path & "\Lap Setoran1.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub Command3_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command3_Click()
    CrystalReport1.SelectionFormula = "ToText({Transaksi.Tanggal})='" & Date & "' and {Transaksi.KodeKsr}='" & LblKodeKsr & "' and {Transaksi.Keterangan}='Pengambilan Kas'"
    CrystalReport1.ReportFileName = App.Path & "\Lap Pengambilan1.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

