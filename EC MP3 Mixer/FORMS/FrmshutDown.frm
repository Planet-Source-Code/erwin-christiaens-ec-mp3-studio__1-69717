VERSION 5.00
Begin VB.Form FrmshutDown 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmshutDown.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   1815
      Picture         =   "FrmshutDown.frx":15C4E
      ScaleHeight     =   330
      ScaleWidth      =   900
      TabIndex        =   6
      Top             =   2280
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   705
      Picture         =   "FrmshutDown.frx":16C08
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   5
      Top             =   2295
      Width           =   885
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   1800
      Picture         =   "FrmshutDown.frx":17BC2
      ScaleHeight     =   330
      ScaleWidth      =   900
      TabIndex        =   4
      Top             =   1860
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   720
      Picture         =   "FrmshutDown.frx":18B7C
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   3
      Top             =   1860
      Width           =   885
   End
   Begin VB.PictureBox PicNee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2340
      Picture         =   "FrmshutDown.frx":19B36
      ScaleHeight     =   330
      ScaleWidth      =   900
      TabIndex        =   2
      Top             =   1005
      Width           =   900
   End
   Begin VB.PictureBox PicJa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1230
      Picture         =   "FrmshutDown.frx":1AAF0
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   1
      Top             =   990
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dit zal Mp3 Stuio afsluiten, wilt u doorgaan?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   450
      Width           =   4275
   End
End
Attribute VB_Name = "FrmshutDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean


Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then ok = True:  Me.Hide
End Sub

Private Sub Form_Load()

    ok = False
    
End Sub

Private Sub PicJa_Click()
  ok = True
  Me.Hide
End Sub

Private Sub PicJa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicJa.Picture = Picture1(1).Image
End Sub

Private Sub PicJa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicJa.Picture = Picture1(0).Image
End Sub

Private Sub PicNee_Click()
  ok = False
  Me.Hide
End Sub

Private Sub PicNee_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicNee.Picture = Picture2(1).Image
End Sub

Private Sub PicNee_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicNee.Picture = Picture2(0).Image
End Sub
