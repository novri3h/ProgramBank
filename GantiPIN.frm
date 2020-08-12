VERSION 5.00
Begin VB.Form GantiPIN 
   Caption         =   "Ganti PIN"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2730
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
   ScaleHeight     =   1470
   ScaleWidth      =   2730
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1250
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1250
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label LblNorek 
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
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label LblNoKartu 
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Kartu"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Rekening"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pin Baru"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pin Baru"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pin Lama"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "GantiPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Text1.PasswordChar = "X"
Text2.PasswordChar = "X"
Text3.PasswordChar = "X"
Text1.MaxLength = 6
Text2.MaxLength = 6
Text3.MaxLength = 6
Text2.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    If Len(Text1) <> 6 Then
        MsgBox "Nomor Pin Harus 6 digit"
        Text1 = ""
        Text1.SetFocus
        Exit Sub
    End If
    If Text1 <> ATM.TxtPin Then
        MsgBox "PIN salah"
        Text1 = ""
        Text1.SetFocus
    Else
        Call BukaDB
        RSNasabah.Open "select * from nasabah where Pin='" & Text1 & "'", Conn
        If Not RSNasabah.EOF Then
            Text1.Enabled = False
            Text2.Enabled = True
            Text2.SetFocus
            LblNoRek = RSNasabah!norek
            LblNoKartu = RSNasabah!Nokartu
        Else
            MsgBox "Pin Salah"
            Text1 = ""
            Text1.SetFocus
        End If
        Conn.Close
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    If Len(Text2) <> 6 Then
        MsgBox "Nomor Pin Harus 6 digit"
        Text2 = ""
        Text2.SetFocus
        Exit Sub
    Else
        Call BukaDB
        RSNasabah.Open "select * from nasabah where PIN='" & Text2 & "'", Conn
        If Not RSNasabah.EOF Then
            MsgBox "PIN ini sudah ada, ganti nomor lain"
            Text2 = ""
            Text2.SetFocus
            Exit Sub
        End If
        Conn.Close
    End If
    MsgBox "Ketik No PIN yang sama sekali lagi"
    Text2.Enabled = False
    Text3.Enabled = True
    Text3 = ""
    Text3.SetFocus
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    If Len(Text3) <> 6 Then
        MsgBox "Nomor Pin Harus 6 digit"
        Text3 = ""
        Text3.SetFocus
        Exit Sub
    End If
    If Text3 <> Text2 Then
        MsgBox "No PIN tidak sama"
        Text3 = ""
        Text3.SetFocus
    Else
        Pesan = MsgBox("Yakin PIN akan diganti..?", vbYesNo, "Konfirmasi")
        If Pesan = vbYes Then
            Call BukaDB
            Dim Berubah As String
            Berubah = "Update nasabah set PIN='" & Text3 & "' where norek='" & LblNoRek & "' and nokartu='" & LblNoKartu & "'"
            Conn.Execute Berubah
            ATM.TxtPin = Text3
            MsgBox "PIN telah berubah menjadi << " & Text3 & ">>"
            Menu.DT.Refresh
            Unload Me
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub
