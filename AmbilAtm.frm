VERSION 5.00
Begin VB.Form AmbilAtm 
   Caption         =   "Pengambilan ATM"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3645
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
   ScaleHeight     =   1500
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox TxtJumlahAbl 
      Height          =   350
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   2000
   End
   Begin VB.Label Nomor 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   2000
   End
   Begin VB.Label LblNama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   2000
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Nasabah"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Transaksi"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "AmbilAtm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB
RSNasabah.Open "Select * from nasabah where PIN='" & ATM.TxtPin & "' and norek='" & ATM.LblNoRek & "' and Nokartu='" & ATM.LblNoKartu & "'", Conn
If Not RSNasabah.EOF Then LblNama = RSNasabah!namansb
End Sub

Private Sub Form_Activate()
Call BukaDB
RSTransaksi.Open "select * from Transaksi Where NoTransaksi In(Select Max(NoTransaksi)From Transaksi)Order By NoTransaksi Desc", Conn
RSTransaksi.Requery
Call Auto
End Sub

Private Sub TxtJumlahAbl_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    RSNasabah.Open "Select * from nasabah where PIN='" & ATM.TxtPin & "' and norek='" & ATM.LblNoRek & "' and Nokartu='" & ATM.LblNoKartu & "'", Conn
    If Not RSNasabah.EOF Then
        If RSNasabah!Saldo < Val(TxtJumlahAbl) Then
            MsgBox "Dana tidak cukup" & Chr(13) & _
            "Saldo Anda tinggal Rp. " & Format(RSNasabah!Saldo, "###,###,###") & ""
            TxtJumlahAbl = ""
            Exit Sub
        Else
            Pesan = MsgBox("Jumlah Pengambilan :<<" & TxtJumlahAbl & ">> Data sudah benar..?", vbYesNo, "Konfirmasi")
            If Pesan = vbYes Then
                Dim Edit As String
                Edit = "Update nasabah set saldo='" & RSNasabah!Saldo - TxtJumlahAbl & "' where PIN='" & ATM.TxtPin & "' and norek='" & ATM.LblNoRek & "' and Nokartu='" & ATM.LblNoKartu & "'" ', Conn 'where norek='" & lblnorek & "'"
                Conn.Execute Edit
                
                Dim SimpanTransaksi As String
                SimpanTransaksi = "Insert into Transaksi(NoTransaksi,NoRek,Tanggal,Jam,Pengeluaran,Keterangan,NoATM) values " & Chr(13) & _
                "('" & Nomor & "','" & ATM.LblNoRek & "','" & CDate(ATM.LblTanggal) & "','" & ATM.LblJam & "','" & TxtJumlahAbl & "','" & AmbilAtm.Caption & "','" & ATM.LblNoATM & "')"
                Conn.Execute SimpanTransaksi
                RSNasabah.Requery
                
                Unload Me
                Menu.DT.Refresh
                CetakLayar
            End If
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Sub CetakLayar()
Tampilkan.Show
Tampilkan.Font = "Courier New"
Call BukaDB
RSTransaksi.Open "select * from Transaksi Where NoTransaksi In(Select Max(NoTransaksi)From Transaksi )Order By NoTransaksi Desc", Conn
RSNasabah.Open "select * from Nasabah where NoRek='" & RSTransaksi!norek & "'", Conn
RSAtm.Open "select * from ATM where NoATM='" & RSTransaksi!NoATM & "'", Conn
Tampilkan.Print
Tampilkan.Print Tab(5); "Bukti Pengambilan (ATM)"
Tampilkan.Print
Tampilkan.Print Tab(5); "Nomor     :  "; RSTransaksi!NoTransaksi
Tampilkan.Print Tab(5); "Tanggal   :  "; RSTransaksi!tanggal
Tampilkan.Print Tab(5); "Jam       :  "; RSTransaksi!Jam
Tampilkan.Print Tab(5); "ATM       :  "; RSAtm!NamaAtm
Tampilkan.Print Tab(5); "Nasabah   :  "; RSNasabah!namansb
Tampilkan.Print Tab(5); "Jumlah Rp :  "; RKanan(RSTransaksi!Pengeluaran, "###,###,###")
Tampilkan.Print Tab(5); "Saldo  Rp :  "; RKanan(RSNasabah!Saldo, "###,###,###")
End Sub

Private Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Private Sub lblnorek_Change()
Call BukaDB
RSNasabah.Open "select * from nasabah where norek='" & LblNoRek & "'", Conn
If Not RSNasabah.EOF Then LblNama = RSNasabah!namansb
End Sub
