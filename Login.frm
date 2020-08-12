VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login Teller"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
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
   ScaleHeight     =   1275
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodeKsr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Frame Frame3 
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama Kasir"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1000
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1245
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    TxtNamaKsr.MaxLength = 35
    TxtPasswordKsr.MaxLength = 15
'    TxtNamaKsr.PasswordChar = "X"
    TxtPasswordKsr.PasswordChar = "X"
    TxtPasswordKsr.Enabled = False
    TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamaKsr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    
    Call BukaDB
    RSKasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    If RSKasir.EOF Then
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            'End
            Unload Me
        End If
    Else
        TxtNamaKsr.Enabled = False
        TxtPasswordKsr.Enabled = True
        TxtPasswordKsr.SetFocus
    End If
End If
End Sub

Private Sub txtpasswordksr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
Dim KodeKasir As String
Dim NamaKasir As String
If Keyascii = 13 Then
    Call BukaDB
    RSKasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
    If RSKasir.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            'End
            Unload Me
        End If
    Else
        'Unload Me
        Me.Visible = False
        Menu.Show
        Menu.StatusBar1.Panels(1).Text = RSKasir!kodeksr
        KodeKasir = RSKasir!kodeksr
        NamaKasir = RSKasir!NamaKsr
        TxtKodeKsr = KodeKasir
        TxtNamaKsr = NamaKasir
        
        Pengambilan.LblKodeKsr = KodeKasir
        Pengambilan.LblNamaKsr = NamaKasir
        
        Setoran.LblKodeKsr = KodeKasir
        Setoran.LblNamaKsr = NamaKasir
        
        Nasabah.LblKodeKsr = KodeKasir
        Nasabah.LblNamaKsr = NamaKasir
        
        'Laporan.LblKodeKsr = KodeKasir
        'Laporan.LblNamaKsr = NamaKasir
    End If
End If
End Sub

