VERSION 5.00
Begin VB.Form Frmpics 
   Caption         =   "Form6"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3585
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicautoFade 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   3540
      Picture         =   "Frmpics.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   660
      TabIndex        =   21
      Top             =   1545
      Width           =   660
   End
   Begin VB.PictureBox PicautoFade 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3540
      Picture         =   "Frmpics.frx":0A0E
      ScaleHeight     =   285
      ScaleWidth      =   675
      TabIndex        =   20
      Top             =   1155
      Width           =   675
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   3510
      Picture         =   "Frmpics.frx":1468
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   19
      Top             =   630
      Width           =   885
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   3525
      Picture         =   "Frmpics.frx":2422
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   18
      Top             =   270
      Width           =   885
   End
   Begin VB.PictureBox Picannuleren 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   2460
      Picture         =   "Frmpics.frx":33DC
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   17
      Top             =   630
      Width           =   885
   End
   Begin VB.PictureBox Picannuleren 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   2460
      Picture         =   "Frmpics.frx":4396
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   16
      Top             =   270
      Width           =   885
   End
   Begin VB.PictureBox PicFind 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   2250
      Picture         =   "Frmpics.frx":5350
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   15
      Top             =   2895
      Width           =   885
   End
   Begin VB.PictureBox PicFindNext 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   1290
      Picture         =   "Frmpics.frx":630A
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   14
      Top             =   2895
      Width           =   885
   End
   Begin VB.PictureBox PicResetfading 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   195
      Picture         =   "Frmpics.frx":72C4
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   13
      Top             =   2910
      Width           =   885
   End
   Begin VB.PictureBox PicFind 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   2250
      Picture         =   "Frmpics.frx":827E
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   12
      Top             =   2505
      Width           =   885
   End
   Begin VB.PictureBox PicFindNext 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   1290
      Picture         =   "Frmpics.frx":9238
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   11
      Top             =   2505
      Width           =   885
   End
   Begin VB.PictureBox PicResetfading 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   195
      Picture         =   "Frmpics.frx":A1F2
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   10
      Top             =   2520
      Width           =   885
   End
   Begin VB.PictureBox PicSingleplay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   1965
      Picture         =   "Frmpics.frx":B1AC
      ScaleHeight     =   285
      ScaleWidth      =   540
      TabIndex        =   9
      Top             =   2130
      Width           =   540
   End
   Begin VB.PictureBox PicSingleplay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   1950
      Picture         =   "Frmpics.frx":B9F2
      ScaleHeight     =   285
      ScaleWidth      =   540
      TabIndex        =   8
      Top             =   1800
      Width           =   540
   End
   Begin VB.PictureBox PicLoop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   1305
      Picture         =   "Frmpics.frx":C238
      ScaleHeight     =   285
      ScaleWidth      =   540
      TabIndex        =   7
      Top             =   2130
      Width           =   540
   End
   Begin VB.PictureBox PicLoop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   1305
      Picture         =   "Frmpics.frx":CA7E
      ScaleHeight     =   285
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   1815
      Width           =   540
   End
   Begin VB.PictureBox PicPlayerShuffle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   195
      Picture         =   "Frmpics.frx":D2C4
      ScaleHeight     =   285
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   2145
      Width           =   915
   End
   Begin VB.PictureBox PicPlayerShuffle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   195
      Picture         =   "Frmpics.frx":E0AE
      ScaleHeight     =   285
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   1830
      Width           =   915
   End
   Begin VB.PictureBox PicbtnPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   0
      Left            =   1035
      Picture         =   "Frmpics.frx":EE98
      ScaleHeight     =   825
      ScaleWidth      =   990
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicbtnPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   1
      Left            =   1035
      Picture         =   "Frmpics.frx":119D2
      ScaleHeight     =   825
      ScaleWidth      =   990
      TabIndex        =   2
      Top             =   870
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicbtnStop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   0
      Left            =   15
      Picture         =   "Frmpics.frx":1450C
      ScaleHeight     =   825
      ScaleWidth      =   990
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicbtnStop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   1
      Left            =   0
      Picture         =   "Frmpics.frx":17046
      ScaleHeight     =   825
      ScaleWidth      =   990
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "Frmpics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
