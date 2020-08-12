VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JejakTransaksi 
   Caption         =   "Audit"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3870
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
   ScaleHeight     =   1290
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1920
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   1500
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No Rekening"
         Height          =   345
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1140
      End
   End
End
Attribute VB_Name = "JejakTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB
RSNasabah.Open "select * from nasabah", Conn
Do While Not RSNasabah.EOF
    Combo1.AddItem RSNasabah!norek
    RSNasabah.MoveNext
Loop

End Sub

Private Sub Combo1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    If Len(Combo1) < 10 Then
        MsgBox "Nomor Rekening harus 10 digit"
        Combo1.SetFocus
        Exit Sub
    End If
    Call BukaDB
    RSNasabah.Open "select * from nasabah where Norek='" & Combo1 & "'", Conn
    If Not RSNasabah.EOF Then
        CrystalReport1.SelectionFormula = "{Transaksi.Norek}='" & Combo1 & "'"
        CrystalReport1.ReportFileName = App.Path & "\Lap Jejak.rpt"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.Action = 1
    Else
        MsgBox "Nomor Rekeing tidak terdaftar"
        Combo1.SetFocus
        Exit Sub
    End If
End If
End Sub

