VERSION 5.00
Begin VB.Form FrmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FrmAbout.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5490
      Picture         =   "FrmAbout.frx":69F92
      ScaleHeight     =   345
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   3945
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2007"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6135
      TabIndex        =   2
      Top             =   885
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Erwin Christiaens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4545
      TabIndex        =   1
      Top             =   615
      Width           =   2115
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Unload Me
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub
