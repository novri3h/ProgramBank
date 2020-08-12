VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6810
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   5085
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "a"
            Object.ToolTipText     =   "Nasabah"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "b"
            Object.ToolTipText     =   "Setoran"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "c"
            Object.ToolTipText     =   "Pengambilan"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "d"
            Object.ToolTipText     =   "Tutup Rekening"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "e"
            Object.ToolTipText     =   "ATM"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "f"
            Object.ToolTipText     =   "Laporan Data Nasabah"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "g"
            Object.ToolTipText     =   "Laporan Setoran"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "h"
            Object.ToolTipText     =   "Laporan Penarikan"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "i"
            Object.ToolTipText     =   "Jejak Transaksi Model Pertama"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "j"
            Object.ToolTipText     =   "Jejak Transaksi Model Kedua"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k"
            Object.ToolTipText     =   "Keluar"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   4590
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   405
      Left            =   120
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin Crystal.CrystalReport CR 
      Left            =   2160
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Menu.frx":33586
      Height          =   2445
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4313
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Rekeing"
         Caption         =   "Rekening"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nasabah"
         Caption         =   "Nasabah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PIN"
         Caption         =   "PIN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":33597
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":338B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":33BCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":33EE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":341FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":34519
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":34833
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":34B4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":34E67
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":35181
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":3549B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnnasabah 
         Caption         =   "&Nasabah"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnsetoran 
         Caption         =   "&Setoran"
      End
      Begin VB.Menu mnpengambilan 
         Caption         =   "&Pengambilan"
      End
      Begin VB.Menu mntutuprek 
         Caption         =   "&Tutup Rekening"
      End
   End
   Begin VB.Menu mnatm 
      Caption         =   "&ATM"
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "&Laporan Hari Ini"
      Begin VB.Menu mnlapnasabahbaru 
         Caption         =   "&Nasabah Baru"
      End
      Begin VB.Menu mnlapsetoranhariini 
         Caption         =   "&Setoran"
      End
      Begin VB.Menu lappenarikanhariini 
         Caption         =   "&Penarikan"
      End
   End
   Begin VB.Menu mnrincian 
      Caption         =   "&Rincian Transaksi"
      Begin VB.Menu lapjejak1 
         Caption         =   "&Jejak Transaksi 1"
      End
      Begin VB.Menu mnjejak 
         Caption         =   "&Jejak Transaksi 2"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DT.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
DT.RecordSource = "select norek as Rekeing,namansb as Nasabah,PIN,Saldo from nasabah"
DT.Refresh
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Pesan = MsgBox("Yakin akan keluar dari program..?", vbYesNo, "Konfirmasi")
    If Pesan = vbYes Then End
End If
End Sub

Private Sub Form_Load()
DT.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
DT.RecordSource = "select norek as Rekeing,namansb as Nasabah,PIN,Saldo from nasabah"
DT.Refresh
Set DataGrid1.DataSource = DT
DataGrid1.Refresh

End Sub

Private Sub lapjejak1_Click()
JejakTransaksi.Show
End Sub

Private Sub lappenarikanhariini_Click()
    CR.SelectionFormula = "{Transaksi.Keterangan}='Pengambilan Kas' and {Transaksi.Tanggal}=today"
    CR.ReportFileName = App.Path & "\Lap Pengambilan1.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnatm_Click()
ATM.Show
End Sub

Private Sub mnjejak_Click()
Laporan2.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnlapnasabahbaru_Click()
    CR.ReportFileName = App.Path & "\Lap Nasabah Baru.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub



Private Sub mnlogin_Click()
Login.Show
End Sub

Private Sub mnlapsetoranhariini_Click()
    CR.SelectionFormula = "{Transaksi.Keterangan}='Setoran Kas' and {Transaksi.Tanggal}=today"
    CR.ReportFileName = App.Path & "\Lap Setoran2.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnnasabah_Click()
Nasabah.Show
End Sub

Private Sub mnpengambilan_Click()
Pengambilan.Show
End Sub

Private Sub mnsetoran_Click()
Setoran.Show
End Sub

Private Sub mntelleratm_Click()
Rincian.Show
End Sub

Private Sub mntutuprek_Click()
Dim Rekening As String
Rekening = InputBox("Nomor Rekening :")
Call BukaDB
RSNasabah.Open "Select * from nasabah where Norek='" & Rekening & "'", Conn
If Not RSNasabah.EOF Then
    Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
    If Pesan = vbYes Then
        Dim HapusNsb As String
        HapusNsb = "delete * from nasabah where norek='" & Rekening & "'"
        Conn.Execute HapusNsb
        RSNasabah.Requery
        Form_Activate
    End If
Else
    MsgBox "Nomor Rekening tidak terdaftar"
    Exit Sub
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Nasabah.Show
    Case "b"
        Setoran.Show
    Case "c"
        Pengambilan.Show
    Case "d"
        Dim Rekening As String
        Rekening = InputBox("Nomor Rekening :")
        Call BukaDB
        RSNasabah.Open "Select * from nasabah where Norek='" & Rekening & "'", Conn
        If Not RSNasabah.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim HapusNsb As String
                HapusNsb = "delete * from nasabah where norek='" & Rekening & "'"
                Conn.Execute HapusNsb
                RSNasabah.Requery
                Form_Activate
            End If
        Else
            MsgBox "Nomor Rekening tidak terdaftar"
            Exit Sub
        End If

    Case "e"
        ATM.Show
    Case "f"
        CR.ReportFileName = App.Path & "\Lap Nasabah Baru.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1

    Case "g"
        CR.SelectionFormula = "{Transaksi.Keterangan}='Setoran Kas' and {Transaksi.Tanggal}=today"
        CR.ReportFileName = App.Path & "\Lap Setoran2.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1

    Case "h"
        CR.SelectionFormula = "{Transaksi.Keterangan}='Pengambilan Kas' and {Transaksi.Tanggal}=today"
        CR.ReportFileName = App.Path & "\Lap Pengambilan1.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "i"
        JejakTransaksi.Show
    Case "j"
        Laporan2.Show
    Case "k"
        End
        
End Select
End Sub
