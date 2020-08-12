VERSION 5.00
Begin VB.Form Ambil1 
   Caption         =   "Pengambilan"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2670
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
   LockControls    =   -1  'True
   ScaleHeight     =   1320
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Pengambilan Manual"
      Height          =   500
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pengambilan Otomatis"
      Height          =   500
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "Ambil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Ambil2.Show
End Sub

Private Sub Command1_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command2_Click()
AmbilAtm.Show
End Sub

Private Sub Command2_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub
