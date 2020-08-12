VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Pembayaran 
   Caption         =   "Bayar Telepon"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
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
   ScaleHeight     =   2865
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNoTujuan 
      Height          =   350
      Left            =   1560
      TabIndex        =   5
      Text            =   " Lihat No. Rek PT Telkom"
      Top             =   1560
      Width           =   2000
   End
   Begin VB.TextBox TxtNomorPlg 
      Height          =   350
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   375
      Left            =   1560
      Top             =   2400
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.Label NamaPrsh 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label Tagihan 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label NamaPlg 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   1995
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Perusahaan"
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
      TabIndex        =   7
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Rekening Tujuan"
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Transaksi"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Pelanggan"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tagihan"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1500
   End
End
Attribute VB_Name = "Pembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DT.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
DT.RecordSource = "Transaksi"
Call Auto
End Sub

Private Sub TxtNomorPlg_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me: ATM.Show
If Keyascii = 13 Then
    Call BukaDB
    RSTagihan.Open "select * from Tagihan where NomorPlg='" & TxtNomorPlg & "'", Conn
    If Not RSTagihan.EOF Then
        NamaPlg = RSTagihan!NamaPlg
        Tagihan = RSTagihan!Tagihan
        If RSTagihan!Status = "Lunas" Then
            Tagihan = ""
            MsgBox "Tagihan telah lunas..!"
            Unload Me
            ATM.Show
            Exit Sub
        Else
            'TxtNoTujuan = ""
            TxtNoTujuan.SetFocus
        End If
    Else
        MsgBox "Nomor pelanggan salah..!"
        TxtNomorPlg = ""
        TxtNomorPlg.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TxtNoTujuan_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me: ATM.Show
If Keyascii = 13 Then
    Call BukaDB
    RSNasabah.Open "select * from Nasabah where NoRek='" & TxtNoTujuan & "'", Conn
    If Not RSNasabah.EOF Then
        NamaPrsh = RSNasabah!namansb
        Pesan = MsgBox("Lakukan pembayaran ke " & NamaPrsh & "..?", vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Call BukaDB
            RSNasabah.Open "select * from nasabah where Pin='" & ATM.TxtPin & "' and Norek='" & ATM.LblNoRek & "' and NoKartu='" & ATM.LblNoKartu & "'", Conn
            If Not RSNasabah.EOF Then
                If RSNasabah!Saldo < Val(Tagihan) Then
                    MsgBox "Dana Anda tidak cukup" & Chr(13) & _
                    "Saldo Anda hanya Rp. " & Format(RSNasabah!Saldo, "###,###,###") & ""
                    Unload Me
                    ATM.Show
                    Exit Sub
                Else
                    Dim EditDanaNsb As String
                    EditDanaNsb = "Update nasabah set saldo='" & RSNasabah!Saldo - Tagihan & "' where Pin='" & ATM.TxtPin & "' and Norek='" & ATM.LblNoRek & "' and NoKartu='" & ATM.LblNoKartu & "'"
                    Conn.Execute EditDanaNsb
                End If
            End If

            Call BukaDB
            RSNasabah.Open "select * from nasabah where Norek='" & TxtNoTujuan & "'", Conn
            Dim TambahSaldoTelkom As String
            TambahSaldoTelkom = "Update nasabah set saldo='" & RSNasabah!Saldo + Tagihan & "' where norek='" & TxtNoTujuan & "' "
            Conn.Execute TambahSaldoTelkom
            
            Dim EditDataTgh As String
            EditDataTgh = "Update tagihan set Status='Lunas' where NomorPlg='" & TxtNomorPlg & "'"
            Conn.Execute EditDataTgh
            
            Call BukaDB
            RSTransaksi.Open "Transaksi", Conn
            Dim SimpanTransaksi1 As String
            SimpanTransaksi1 = "Insert into Transaksi(NoTransaksi,NoRek,Tanggal,Jam,Pengeluaran,Keterangan,NomorPlg,NoRekTjn,NoATM) values " & Chr(13) & _
            "('" & Nomor & "','" & ATM.LblNoRek & "','" & CDate(ATM.LblTanggal) & "','" & ATM.LblJam & "','" & Tagihan & "','" & Pembayaran.Caption & "','" & TxtNomorPlg & "','" & TxtNoTujuan & "','" & ATM.LblNoATM & "')"
            Conn.Execute SimpanTransaksi1
            Conn.Close
            
            Call BukaDB
            RSTransaksi.Open "Transaksi", Conn
            Dim SimpanTransaksi2 As String
            SimpanTransaksi2 = "insert into Transaksi(NoTransaksi,NoRek,Tanggal,Jam,Pemasukan,Keterangan,NoATM) values " & Chr(13) & _
            "('" & Nomor + 1 & "','" & TxtNoTujuan & "','" & ATM.LblTanggal & "','" & ATM.LblJam & "','" & Tagihan & "','" & Pembayaran.Caption + " Dari " + NamaPlg & "','" & ATM.LblNoATM & "')"
            Conn.Execute SimpanTransaksi2
            Conn.Close
                         
            'Dim HapusTagihan As String
            'HapusTagihan = "delete * from Tagihan where nomorplg='" & TxtNomorPlg & "'"
            'Conn.Execute HapusTagihan
            
            Form_Activate
            Unload Me
            CetakLayar
        Else
            Unload Me
            Exit Sub
        End If
    Else
        MsgBox "Nomor rekening Tujuan tidak terdaftar"
        TxtNoTujuan.SetFocus
        Exit Sub
    End If
End If
Menu.DT.Refresh
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
RSTagihan.Open "select * from Tagihan where NomorPlg='" & DT.Recordset!Nomorplg & "'", Conn
Tampilkan.Print
Tampilkan.Print Tab(5); "Bukti Pembayaran(ATM)"
Tampilkan.Print
Tampilkan.Print Tab(5); "Nomor        :  "; DT.Recordset!NoTransaksi
Tampilkan.Print Tab(5); "Tanggal      :  "; DT.Recordset!tanggal
Tampilkan.Print Tab(5); "Jam          :  "; DT.Recordset!Jam
Tampilkan.Print Tab(5); "No ATM       :  "; DT.Recordset!NoATM
Tampilkan.Print Tab(5); "Dibayar Oleh :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "No Pelanggan :  "; RSTagihan!Nomorplg
Tampilkan.Print Tab(5); "Pelanggan    :  "; RSTagihan!NamaPlg
Conn.Close
Call BukaDB
RSNasabah.Open "select * from Nasabah where NoRek='" & DT.Recordset!NoRekTjn & "'", Conn
Tampilkan.Print Tab(5); "Penerima     :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "Jumlah       :  "; Format(DT.Recordset!Pengeluaran, "###,###,###")
End Sub

Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

