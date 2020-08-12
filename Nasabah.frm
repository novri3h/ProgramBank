VERSION 5.00
Begin VB.Form Nasabah 
   Caption         =   "Buka Rekening"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3360
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2760
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1500
   End
   Begin VB.TextBox TxtNoRek 
      Height          =   350
      Left            =   1440
      TabIndex        =   23
      Top             =   1440
      Width           =   1250
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   120
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   4560
      TabIndex        =   3
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2880
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   1200
   End
   Begin VB.TextBox TxtNama 
      Height          =   350
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Width           =   4300
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label LblSaldo 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   22
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label LblPIN 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   21
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label LblNoKartu 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   20
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label LblJam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   19
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label LblKodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   18
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   17
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   16
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label LblNamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   15
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
      Left            =   3600
      TabIndex        =   14
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Transaksi"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      Height          =   345
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Kartu"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PIN"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Nasabah"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Rekeing"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1245
   End
End
Attribute VB_Name = "Nasabah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
LblTanggal = Date
LblKodeKsr = Login.TxtKodeKsr
LblNamaKsr = Login.TxtNamaKsr
TxtNama.MaxLength = 30
KondisiAwal
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

Call AcakNorek
Call Auto
Nomor = Nomor
LblSaldo = 500000
End Sub

Private Sub Timer1_Timer()
LblJam = Time$
End Sub

Function CariData()
Call BukaDB
RSNasabah.Open "Select * From Nasabah where NoRek='" & TxtNoRek & "'", Conn
End Function

Function CariCombo()
Call BukaDB
RSNasabah.Open "Select * From Nasabah where NoRek='" & Combo1 & "'", Conn
End Function

Private Sub KosongkanText()
TxtNoRek = "": TxtNama = ""
End Sub

Private Sub SiapIsi()
TxtNama.Enabled = True
End Sub

Private Sub TidakSiapIsi()
TxtNama.Enabled = False
End Sub

Private Sub KondisiAwal()
Form_Activate
KosongkanText
TidakSiapIsi
CmdInput.Caption = "&Input"
CmdEdit.Caption = "&Edit"
CmdHapus.Caption = "&Hapus"
CmdTutup.Caption = "&Tutup"
CmdInput.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
With RSNasabah
If Not RSNasabah.EOF Then
    TxtNoRek = RSNasabah!norek
    TxtNama = RSNasabah!namansb
    LblSaldo = RSNasabah!Saldo
    LblNoKartu = RSNasabah!Nokartu
    LblPIN = RSNasabah!PIN
End If
End With
End Sub

Private Sub CmdInput_Click()
If CmdInput.Caption = "&Input" Then
    CmdInput.Caption = "&Simpan"
    CmdEdit.Enabled = False
    CmdHapus.Enabled = False
    CmdTutup.Caption = "&Batal"
    SiapIsi
    KosongkanText
    AcakNorek
    AcakPinKartu
    TxtNoRek.Enabled = False
    TxtNama.SetFocus
Else
    If TxtNoRek = "" Or TxtNama = "" Then
        MsgBox "Data Belum Lengkap...!"
    Else
        Call CariData
        If Not RSNasabah.EOF Then
            TxtNama.Enabled = False
            CmdInput.SetFocus
            Exit Sub
        Else
            Dim SimpanNasabah As String
            SimpanNasabah = "Insert Into Nasabah (NoRek,NamaNsb,Saldo,NoKartu,PIN,Tgldaftar) values " & Chr(13) & _
            "('" & TxtNoRek & "','" & TxtNama & "','" & LblSaldo & "','" & LblNoKartu & "','" & LblPIN & "','" & LblTanggal & "')"
            Conn.Execute SimpanNasabah
            
            Dim SimpanTransaksi As String
            SimpanTransaksi = "Insert into Transaksi(NoTransaksi,Norek,Tanggal,Jam,Pemasukan,Keterangan,KodeKsr) values " & Chr(13) & _
            "('" & Nomor & "','" & TxtNoRek & "','" & CDate(LblTanggal) & "','" & LblJam & "','" & Val(LblSaldo) & "','" & Nasabah.Caption & "','" & Menu.StatusBar1.Panels(1).Text & "')"
            Conn.Execute SimpanTransaksi
            
            RefreshData
            Form_Activate
            KondisiAwal
            Menu.DT.Refresh
            CetakLayar
        End If
    End If
End If
End Sub

Private Sub CmdEdit_Click()
If CmdEdit.Caption = "&Edit" Then
    CmdInput.Enabled = False
    CmdEdit.Caption = "&Simpan"
    CmdHapus.Enabled = False
    CmdTutup.Caption = "&Batal"
    SiapIsi
    TxtNoRek.Enabled = False
    TxtNoRek = ""
    Combo1.SetFocus
Else
    If TxtNama = "" Then
        MsgBox "Masih Ada Data Yang Kosong"
    Else
        Dim EditNasabah As String
        EditNasabah = "Update Nasabah Set NamaNsb= '" & TxtNama & "' where NoRek='" & Combo1 & "'"
        Conn.Execute EditNasabah
        RefreshData
        KondisiAwal
        Menu.DT.Refresh
    End If
End If
End Sub

Private Sub CmdHapus_Click()
If CmdHapus.Caption = "&Hapus" Then
    CmdHapus.Caption = "&Delete"
    CmdInput.Enabled = False
    CmdEdit.Enabled = False
    CmdTutup.Caption = "&Batal"
    KosongkanText
    SiapIsi
    TxtNoRek.Enabled = False
    Combo1.SetFocus
End If
End Sub

Private Sub CmdTutup_Click()
Select Case CmdTutup.Caption
    Case "&Tutup"
        RefreshData
        Unload Me
    Case "&Batal"
        TidakSiapIsi
        KondisiAwal
        RefreshData
    End Select
End Sub

Private Sub Combo1_Click()
Call BukaDB
RSNasabah.Open "select * from nasabah where norek='" & Combo1 & "'", Conn
If Not RSNasabah.EOF Then
    With RSNasabah
    If Not RSNasabah.EOF Then
        TxtNoRek = RSNasabah!norek
        TxtNama = RSNasabah!namansb
        LblSaldo = RSNasabah!Saldo
        LblNoKartu = RSNasabah!Nokartu
        LblPIN = RSNasabah!PIN
    End If
    End With
End If
End Sub

Private Sub Combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSNasabah.EOF Then
                TampilkanData
                TxtNoRek.Enabled = False
                TxtNama.SetFocus
            Else
                MsgBox "Nomor Rekening Nasabah Tidak Ada"
                TxtNoRek.SetFocus
                Exit Sub
            End If
    End If
    
    If CmdHapus.Caption = "&Delete" Then
        Call CariData
            If Not RSNasabah.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim HapusNasabah As String
                    HapusNasabah = "Delete From Nasabah where NoRek= '" & TxtNoRek & "'"
                    Conn.Execute HapusNasabah
                    KondisiAwal
                    Menu.DT.Refresh
                Else
                    KondisiAwal
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                TxtNoRek.SetFocus
            End If
    End If
End If
End Sub

Private Sub TxtNoRek_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(TxtNoRek) < 10 Then
        MsgBox "Nomor Rekening Harus 10 Digit"
        TxtNoRek.SetFocus
        Exit Sub
    Else
        TxtNama.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSNasabah.EOF Then
                TampilkanData
                MsgBox "Nomor Rekening Sudah Ada"
                KosongkanText
                TxtNoRek.SetFocus
                Exit Sub
            Else
                TxtNama.SetFocus
            End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TxtNama_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If CmdInput.Enabled = True Then
        CmdInput.SetFocus
    ElseIf CmdEdit.Enabled = True Then
        CmdEdit.SetFocus
    End If
End If
End Sub

Sub AcakNorek()
Randomize
MyValue = Int((Val(Nomor) * Rnd()) + 1234567891)
TxtNoRek = MyValue
End Sub

Sub AcakPinKartu()
Randomize
MyValue = Int((Val(TxtNoRek) * Rnd()) + 1234567891)
LblNoKartu = MyValue

Randomize
MyValue = Int((Val(Mid(TxtNoRek, 2, 6)) * Rnd()) + 123456)
LblPIN = MyValue
End Sub

Sub RefreshData()
Call BukaDB
RSNasabah.Open "nasabah", Conn
RSNasabah.Requery
End Sub

Sub CetakLayar()
Tampilkan.Show
Tampilkan.Font = "Courier New"
Call BukaDB
RSTransaksi.Open "select * from Transaksi Where NoTransaksi In(Select Max(NoTransaksi)From Transaksi)Order By NoTransaksi Desc", Conn
RSNasabah.Open "select * from Nasabah where NoRek='" & RSTransaksi!norek & "'", Conn
RSKasir.Open "select * from Kasir where KodeKsr='" & RSTransaksi!kodeksr & "'", Conn
Tampilkan.Print
Tampilkan.Print Tab(5); "Buka Rekening (Teller)"
Tampilkan.Print
Tampilkan.Print Tab(5); "Nomor      :  "; RSTransaksi!NoTransaksi
Tampilkan.Print Tab(5); "Tanggal    :  "; RSTransaksi!tanggal
Tampilkan.Print Tab(5); "Jam        :  "; RSTransaksi!Jam
Tampilkan.Print Tab(5); "Kasir      :  "; RSKasir!NamaKsr
Tampilkan.Print Tab(5); "Nasabah    :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "Jumlah Rp  :  "; RSTransaksi!Pemasukan
Tampilkan.Print Tab(5); "Saldo  Rp  :  "; RSNasabah!Saldo
End Sub

Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

