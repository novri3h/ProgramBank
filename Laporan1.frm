VERSION 5.00
Begin VB.Form Laporan1 
   Caption         =   "Laporan Jejak Transaksi"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   960
      Width           =   2500
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   600
      Width           =   2500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Transaksi"
      Height          =   350
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Rekening"
      Height          =   350
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "Laporan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Call BukaDB
RSTransaksi.Open "select distinct keterangan from transaksi where norek='" & Combo1 & "'", Conn
Do While Not RSTransaksi.EOF
    Combo2.AddItem RSTransaksi!keterangan
    RSTransaksi.MoveNext
Loop
Conn.Close
End Sub

Private Sub Combo2_Click()
Call BukaDB
RSTransaksi.Open "select distinct tanggal from transaksi where norek='" & Combo1 & "'", Conn
Do While Not RSTransaksi.EOF
    Combo3.AddItem RSTransaksi!tanggal
    RSTransaksi.MoveNext
Loop
Conn.Close

End Sub

Private Sub Form_Load()
Call BukaDB
RSNasabah.Open "select * from nasabah", Conn
Do While Not RSNasabah.EOF
    Combo1.AddItem RSNasabah!norek
    RSNasabah.MoveNext
Loop
Conn.Close



'Combo2.AddItem "BUKA REKENING"
'Combo2.AddItem "SETORAN KAS"
'Combo2.AddItem "PENGAMBILAN KAS"
'Combo2.AddItem "PENGAMBILAN ATM"
'Combo2.AddItem "BAYAR TELEPON"
'Combo2.AddItem "TRANSFER DANA"
End Sub
