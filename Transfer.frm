VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Transfer 
   Caption         =   "Transfer Dana"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3660
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
   ScaleHeight     =   2760
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc DT 
      Height          =   375
      Left            =   240
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtJumlahTsf 
      Height          =   350
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox TxtNoTujuan 
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   2000
   End
   Begin VB.Label Penerima 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   9
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label Pengirim 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   2000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pengirim"
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Transaksi"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Transfer"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Penerima"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Tujuan"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1395
   End
End
Attribute VB_Name = "Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DT.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
DT.RecordSource = "Transaksi"
Call BukaDB
RSTransaksi.Open "select * from Transaksi Where NoTransaksi In(Select Max(NoTransaksi)From Transaksi)Order By NoTransaksi Desc", Conn
RSTransaksi.Requery
Call Auto
End Sub

Private Sub TxtNoTujuan_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    RSNasabah.Open "select * from nasabah where norek='" & TxtNoTujuan & "'", Conn
    If RSNasabah.EOF Then
        MsgBox "No Rekening salah"
    Else
        Penerima = RSNasabah!namansb
        TxtJumlahTsf.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TxtJumlahTsf_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    RSNasabah.Open "select * from nasabah where pin='" & ATM.TxtPin & "' and Norek='" & ATM.LblNoRek & "' and NoKartu='" & ATM.LblNoKartu & "'", Conn
    If RSNasabah!Saldo < Val(TxtJumlahTsf) Then
        MsgBox "Dana tidak cukup" & Chr(13) & _
        "Saldo hanya Rp. " & Format(RSNasabah!Saldo, "###,###,###") & ""
        TxtJumlahTsf = ""
        TxtJumlahTsf.SetFocus
        Exit Sub
        Conn.Close
    Else
        Pesan = MsgBox("Data sudah benar", vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Call BukaDB
            RSNasabah.Open "select * from nasabah where pin='" & ATM.TxtPin & "' and Norek='" & ATM.LblNoRek & "' and NoKartu='" & ATM.LblNoKartu & "'", Conn
            Dim KurangiSaldoPengirim As String
            KurangiSaldoPengirim = "Update nasabah set saldo=" & RSNasabah!Saldo - TxtJumlahTsf & " where pin='" & ATM.TxtPin & "' and Norek='" & ATM.LblNoRek & "' and NoKartu='" & ATM.LblNoKartu & "'"
            Conn.Execute KurangiSaldoPengirim
            Conn.Close
                                
            Call BukaDB
            RSNasabah.Open "select * from nasabah where norek='" & TxtNoTujuan & "'", Conn
            Dim TambahSaldoPenerima As String
            TambahSaldoPenerima = "Update nasabah set saldo=" & RSNasabah!Saldo + TxtJumlahTsf & " where norek='" & TxtNoTujuan & "'"
            Conn.Execute TambahSaldoPenerima
            Conn.Close
            
            Call BukaDB
            Dim SimpanTransaksi1 As String
            SimpanTransaksi1 = "insert into Transaksi(NoTransaksi,NoRek,Tanggal,Jam,Pengeluaran,Keterangan,NorekTjn,NoATM) values " & Chr(13) & _
            "('" & Nomor & "','" & ATM.LblNoRek & "','" & ATM.LblTanggal & "','" & ATM.LblJam & "','" & TxtJumlahTsf & "','" & Transfer.Caption + " Ke " + Penerima & "','" & TxtNoTujuan & "','" & ATM.LblNoATM & "')"
            Conn.Execute SimpanTransaksi1
            Conn.Close
            
            Call BukaDB
            Dim SimpanTransaksi2 As String
            SimpanTransaksi2 = "insert into Transaksi(NoTransaksi,NoRek,Tanggal,Jam,Pemasukan,Keterangan,NoATM) values " & Chr(13) & _
            "('" & Nomor + 1 & "','" & TxtNoTujuan & "','" & ATM.LblTanggal & "','" & ATM.LblJam & "','" & TxtJumlahTsf & "','" & Transfer.Caption + " Dari " + Pengirim & "','" & ATM.LblNoATM & "')"
            Conn.Execute SimpanTransaksi2
            Conn.Close
            
            Unload Me
            Menu.DT.Refresh
            CetakLayar
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Sub CetakLayar()
Tampilkan.Show
Tampilkan.Font = "Courier New"
Call BukaDB
DT.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
DT.RecordSource = "Transaksi"
DT.Refresh
DT.Recordset.MoveLast
DT.Recordset.MovePrevious
RSNasabah.Open "select * from Nasabah where NoRek='" & DT.Recordset!norek & "'", Conn
Tampilkan.Print
Tampilkan.Print Tab(5); "Bukti Transfer (ATM)"
Tampilkan.Print
Tampilkan.Print Tab(5); "Nomor    :  "; DT.Recordset!NoTransaksi
Tampilkan.Print Tab(5); "Tanggal  :  "; DT.Recordset!tanggal
Tampilkan.Print Tab(5); "Jam      :  "; DT.Recordset!Jam
Tampilkan.Print Tab(5); "ATM      :  "; DT.Recordset!NoATM
Tampilkan.Print Tab(5); "Pengirim :  "; RSNasabah!namansb
Conn.Close
Call BukaDB
RSNasabah.Open "select * from Nasabah where NoRek='" & DT.Recordset!NoRekTjn & "'", Conn
Tampilkan.Print Tab(5); "Penerima :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "Jumlah   :  "; Format(DT.Recordset!Pengeluaran, "###,###,###")
End Sub

Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

