VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Kasir 
   Caption         =   "Pengolahan Data Kasir"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNama 
      DataField       =   "NamaKsr"
      DataSource      =   "Adodc1"
      Height          =   400
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   3795
   End
   Begin VB.TextBox TxtPassword 
      DataField       =   "PasswordKsr"
      DataSource      =   "Adodc1"
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "X"
      TabIndex        =   6
      Top             =   1080
      Width           =   3795
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   750
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   750
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   750
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   750
   End
   Begin VB.CommandButton CmdAwal 
      Caption         =   "||<<"
      Height          =   400
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   500
   End
   Begin VB.CommandButton CmdMundur 
      Caption         =   "<<"
      Height          =   400
      Left            =   3600
      TabIndex        =   8
      Top             =   1680
      Width           =   500
   End
   Begin VB.CommandButton CmdMaju 
      Caption         =   ">>"
      Height          =   400
      Left            =   4080
      TabIndex        =   9
      Top             =   1680
      Width           =   500
   End
   Begin VB.CommandButton CmdAkhir 
      Caption         =   ">>||"
      Height          =   400
      Left            =   4560
      TabIndex        =   10
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox TxtKode 
      DataField       =   "KodeKsr"
      DataSource      =   "Adodc1"
      Height          =   400
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   3240
      Top             =   120
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Koleksi Program VB\Program Bank\DBBank.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Koleksi Program VB\Program Bank\DBBank.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Kasir"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      Height          =   400
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Kasir"
      Height          =   405
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Password"
      Height          =   405
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1200
   End
End
Attribute VB_Name = "Kasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub form_load()
'buka database
Call BukaDB
'batasi jumlah karakter
TxtKode.MaxLength = 5
TxtNama.MaxLength = 25
TxtPassword.MaxLength = 10
'jalankan prosedur "kondisiawal"
KondisiAwal
End Sub

Private Sub CmdInput_Click()
'jika cmdinput caption-nya input
If CmdInput.Caption = "&Input" Then
    'ubah caption cmdinput jadi simpan
    CmdInput.Caption = "&Simpan"
    'cmdedit dan hapus tidak dapat digunakan
    CmdEdit.Enabled = False
    CmdHapus.Enabled = False
    'caption cmdtutup  jadi batal
    CmdTutup.Caption = "&Batal"
    'tambahkan satu record baru
    Adodc1.Recordset.AddNew
    KosongkanText
    BukaText
    TxtKode.SetFocus
Else
    'mencegah data yang masih kosong
    If TxtKode = "" Or TxtNama = "" Or TxtPassword = "" Then
        MsgBox "Masih ada data yang kosong"
        Exit Sub
    Else
        'jika cmdinput diklik saat captionnya 'simpan'
        'lakukan update tabel
        Adodc1.Recordset.Update
        'kembali ke kondisi awal
        KondisiAwal
    End If
End If
End Sub

Private Sub CmdEdit_Click()
'jika cmdedit diklik dan recordya masih kosong
'tampilkan pesan...
If Adodc1.Recordset.RecordCount = 0 Then
    pesan = MsgBox("Data Kosong")
    Exit Sub
End If
'jika cmdedit caption-nya edit, ubah jadi 'simpan'
If CmdEdit.Caption = "&Edit" Then
    'cmdinput dan hapus tidak dapat digunakan
    CmdInput.Enabled = False
    CmdEdit.Caption = "&Simpan"
    CmdHapus.Enabled = False
    'cmdtutup jadi batal
    CmdTutup.Caption = "&Batal"
    BukaText
    TxtKode.Enabled = False
    TxtNama.SetFocus
Else
    'mencegah nama dan password jika kosong
    If TxtNama = "" Or TxtPassword = "" Then
        MsgBox "masih ada data yang kosong"
        Exit Sub
    Else
        Adodc1.Recordset.Update
        KondisiAwal
    End If
End If
End Sub

Private Sub CmdHapus_Click()
'data tidak dapat dihapus jika kosong
If Adodc1.Recordset.RecordCount = 0 Then
    pesan = MsgBox("Data Kosong")
    Exit Sub
Else
    'jika datanya ada, tampilkan pesan
    pesan = MsgBox("Yakin data ini akan dihapus..?", vbYesNo, "Konfirmasi")
    If pesan = vbYes Then
        'jika dijawab Yes, hapus data tersebut
        Adodc1.Recordset.Delete
        Adodc1.Recordset.Requery
    End If
End If
End Sub

'cmdtutup bekerja sesuai captionnya
Private Sub CmdTutup_Click()
Select Case CmdTutup.Caption
    Case "&Tutup"
        Adodc1.Recordset.CancelBatch
        Unload Me
    Case "&Batal"
        Adodc1.Recordset.CancelBatch
        KondisiAwal
        Adodc1.Refresh
End Select
End Sub

'mencegah entri kode kasir yang sama
Sub TxtKode_Keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    'jika menekan enter setelah mengisi kode kasir,
    'buka database
    Call BukaDB
    'cari kode kasir yang sama
    RSKasir.Open "select * from Kasir where KodeKsr='" & TxtKode & "'", Conn
    RSKasir.Requery
    'jika ditemukan tampilkan datanya dan munculkan pesan...
    If Not RSKasir.EOF Then
        TxtNama = RSKasir!NamaKsr
        TxtPassword = RSKasir!PasswordKsr
        MsgBox "Kode sudah ada, ganti kode lain"
        KosongkanText
        TxtKode.SetFocus
    Else
        'jika tidak ditemukan, lanjutkan mengisi nama
        TxtNama.SetFocus
    End If
End If
End Sub

Sub TxtNama_Keypress(KeyAscii As Integer)
'ubah text jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then TxtPassword.SetFocus
End Sub

Sub TxtPassword_Keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If CmdInput.Caption = "&Simpan" Then
        CmdInput.SetFocus
    ElseIf CmdEdit.Caption = "&Simpan" Then
        CmdEdit.SetFocus
    End If
End If
End Sub

'tombol - tombol navigasi

Private Sub CmdAwal_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data Kosong"
    Exit Sub
Else
    Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub CmdMundur_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data Kosong"
    Exit Sub
Else
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveFirst
        MsgBox "Ini data pertama"
    End If
End If
End Sub

Private Sub CmdMaju_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data Kosong"
    Exit Sub
Else
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveLast
        MsgBox "Ini data terakhir"
    End If
End If
End Sub

Private Sub CmdAkhir_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data Kosong"
    Exit Sub
Else
    Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub KosongkanText()
TxtKode = ""
TxtNama = ""
TxtPassword = ""
End Sub

Sub KondisiAwal()
KunciText
CmdInput.Caption = "&Input"
CmdEdit.Caption = "&Edit"
CmdTutup.Caption = "&Tutup"
CmdInput.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
End Sub

Sub KunciText()
TxtKode.Enabled = False
TxtNama.Enabled = False
TxtPassword.Enabled = False
End Sub

Sub BukaText()
TxtKode.Enabled = True
TxtNama.Enabled = True
TxtPassword.Enabled = True
End Sub

Public Sub Form_Unload(cancel As Integer)
HidupMenu
End Sub
