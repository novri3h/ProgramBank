VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Laporan2 
   Caption         =   "Jejak Transaksi"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
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
   ScaleHeight     =   3600
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   345
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   1600
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Laporan2.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Keterangan"
         Caption         =   "Keterangan"
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
         DataField       =   "Pemasukan"
         Caption         =   "Pemasukan"
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
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   6240
      Top             =   120
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
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
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1600
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1600
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   345
      Left            =   6240
      Top             =   480
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
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
      Caption         =   "Adodc2"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Laporan2.frx":0015
      Height          =   2535
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Keterangan"
         Caption         =   "Keterangan"
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
         DataField       =   "Pengeluaran"
         Caption         =   "Pengeluaran"
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
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " s/d Tanggal"
      Height          =   345
      Left            =   3120
      TabIndex        =   7
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1250
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Rekening"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1250
   End
End
Attribute VB_Name = "Laporan2"
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
Conn.Close
End Sub

Private Sub Combo1_Click()
Call BukaDB
RSTransaksi.Open "select distinct tanggal from transaksi where norek='" & Combo1 & "'", Conn
Combo2.Clear
Do While Not RSTransaksi.EOF
    Combo2.AddItem RSTransaksi!tanggal
    RSTransaksi.MoveNext
Loop
RSNasabah.Open "Select namansb from nasabah where norek='" & Combo1 & "'", Conn
Label4 = RSNasabah!namansb
Conn.Close
End Sub

Private Sub Combo2_Click()
Call BukaDB
RSTransaksi.Open "select distinct tanggal from transaksi where norek='" & Combo1 & "'", Conn
Combo3.Clear
Do While Not RSTransaksi.EOF
    Combo3.AddItem RSTransaksi!tanggal
    RSTransaksi.MoveNext
Loop
Conn.Close

Adodc1.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
Adodc1.RecordSource = "select Keterangan,Pemasukan from transaksi where norek='" & Combo1 & "' and cdate(tanggal)='" & Combo2 & "' and pemasukan<>0"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Adodc2.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
Adodc2.RecordSource = "select Keterangan, Pengeluaran from transaksi where norek='" & Combo1 & "' and cdate(tanggal)='" & Combo2 & "' and pengeluaran<>0"
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2
DataGrid1.Refresh
End Sub

Private Sub Combo3_Click()
Adodc1.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
Adodc1.RecordSource = "select Keterangan,Pemasukan from transaksi where norek='" & Combo1 & "' and cdate(tanggal)>='" & Combo2 & "' and cdate(tanggal)<='" & Combo3 & "' and pemasukan<>0"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Adodc2.ConnectionString = Lokasi '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOBank.mdb"
Adodc2.RecordSource = "select Keterangan, Pengeluaran from transaksi where norek='" & Combo1 & "' and cdate(tanggal)>='" & Combo2 & "' and cdate(tanggal)<='" & Combo3 & "' and pengeluaran<>0"
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2
DataGrid1.Refresh
End Sub

