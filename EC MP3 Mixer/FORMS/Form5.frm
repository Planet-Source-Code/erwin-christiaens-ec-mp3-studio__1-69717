VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   10665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14445
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   14445
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicAutoFade 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6300
      Picture         =   "Form5.frx":08CA
      ScaleHeight     =   285
      ScaleWidth      =   675
      TabIndex        =   198
      Top             =   7260
      Width           =   675
   End
   Begin ECMixer.ECScrollingText ECScrollingText1 
      Height          =   285
      Left            =   1305
      TabIndex        =   196
      Top             =   1395
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   503
   End
   Begin VB.PictureBox bar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      Picture         =   "Form5.frx":1324
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   964
      TabIndex        =   193
      Top             =   0
      Width           =   14460
      Begin VB.Image Image1 
         Height          =   165
         Left            =   150
         Picture         =   "Form5.frx":F556
         Top             =   75
         Width           =   180
      End
      Begin VB.Image Imgmin 
         Height          =   165
         Left            =   13950
         Picture         =   "Form5.frx":F724
         Top             =   75
         Width           =   180
      End
      Begin VB.Image ImgClose 
         Height          =   165
         Left            =   14145
         Picture         =   "Form5.frx":F8F2
         Top             =   75
         Width           =   180
      End
   End
   Begin VB.PictureBox PicFindNext 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   45
      Picture         =   "Form5.frx":FAC0
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   191
      Top             =   7995
      Width           =   885
   End
   Begin VB.PictureBox PicFind 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   960
      Picture         =   "Form5.frx":10A7A
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   190
      Top             =   7995
      Width           =   885
   End
   Begin VB.PictureBox PicResetFader 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7110
      Picture         =   "Form5.frx":11A34
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   189
      Top             =   7245
      Width           =   885
   End
   Begin VB.TextBox TxtFind 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000004&
      Height          =   270
      Left            =   45
      TabIndex        =   185
      Top             =   7650
      Width           =   2985
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   150
      TabIndex        =   180
      Top             =   8775
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   8070
      Picture         =   "Form5.frx":129EE
      ScaleHeight     =   600
      ScaleWidth      =   6300
      TabIndex        =   143
      Top             =   7035
      Width           =   6300
      Begin VB.PictureBox Picture17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5715
         Picture         =   "Form5.frx":1EF10
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   168
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5715
         Picture         =   "Form5.frx":1F3C6
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   167
         Top             =   330
         Width           =   285
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   23
         Left            =   3195
         Picture         =   "Form5.frx":1F87C
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   157
         Top             =   975
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   23
         Left            =   3675
         Picture         =   "Form5.frx":2027E
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   156
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   22
         Left            =   3090
         Picture         =   "Form5.frx":20C80
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   155
         Top             =   105
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   21
         Left            =   2475
         Picture         =   "Form5.frx":21682
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   154
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   20
         Left            =   1875
         Picture         =   "Form5.frx":22084
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   153
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   19
         Left            =   1290
         Picture         =   "Form5.frx":22A86
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   152
         ToolTipText     =   "save playlist"
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   22
         Left            =   2685
         Picture         =   "Form5.frx":23488
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   151
         Top             =   975
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   21
         Left            =   2160
         Picture         =   "Form5.frx":23E8A
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   150
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   20
         Left            =   1650
         Picture         =   "Form5.frx":2488C
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   149
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   19
         Left            =   1125
         Picture         =   "Form5.frx":2528E
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   148
         Top             =   945
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   18
         Left            =   615
         Picture         =   "Form5.frx":25C90
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   147
         Top             =   945
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   18
         Left            =   690
         Picture         =   "Form5.frx":26692
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   146
         ToolTipText     =   "laad playlist"
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   17
         Left            =   75
         Picture         =   "Form5.frx":27094
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   145
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   17
         Left            =   90
         Picture         =   "Form5.frx":27A96
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   144
         ToolTipText     =   "wis lijst"
         Top             =   105
         Width           =   480
      End
      Begin VB.Label LblTotalTimePlaylist2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFE8B6&
         Height          =   210
         Left            =   5280
         TabIndex        =   183
         Top             =   60
         Width           =   315
      End
      Begin VB.Image Image8 
         Height          =   150
         Left            =   4890
         Picture         =   "Form5.frx":28498
         Top             =   375
         Width           =   165
      End
      Begin VB.Image Image7 
         Height          =   150
         Left            =   4695
         Picture         =   "Form5.frx":28642
         Top             =   375
         Width           =   165
      End
      Begin VB.Image Image6 
         Height          =   150
         Left            =   4530
         Picture         =   "Form5.frx":287EC
         Top             =   375
         Width           =   165
      End
      Begin VB.Image Image5 
         Height          =   150
         Left            =   4335
         Picture         =   "Form5.frx":28996
         Top             =   375
         Width           =   165
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   165
         Left            =   5175
         TabIndex        =   158
         ToolTipText     =   "Position of playing track in playlist"
         Top             =   360
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6870
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":28B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":28C12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvPlaylist2 
      Height          =   3195
      Left            =   8070
      TabIndex        =   141
      Top             =   3825
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16771254
      BackColor       =   -2147483642
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Artist"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Album"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Genre"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Track No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tracks Total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Year"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Duration"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Bit Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Comments"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   7335
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   137
      Top             =   6675
      Visible         =   0   'False
      Width           =   6480
   End
   Begin VB.CheckBox ChkAutomaticMixing 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Automatic Mixing"
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   7065
      TabIndex        =   115
      Top             =   6945
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox PicFadeA 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6300
      Picture         =   "Form5.frx":28CE4
      ScaleHeight     =   285
      ScaleWidth      =   660
      TabIndex        =   114
      ToolTipText     =   "start fading"
      Top             =   6900
      Width           =   660
   End
   Begin VB.PictureBox PicFadeB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7365
      Picture         =   "Form5.frx":296F2
      ScaleHeight     =   285
      ScaleWidth      =   660
      TabIndex        =   113
      ToolTipText     =   "start fading"
      Top             =   6900
      Width           =   660
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   6330
      Picture         =   "Form5.frx":2A100
      ScaleHeight     =   3045
      ScaleWidth      =   1755
      TabIndex        =   107
      Top             =   3825
      Width           =   1755
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   90
         Picture         =   "Form5.frx":3B862
         ScaleHeight     =   705
         ScaleWidth      =   1560
         TabIndex        =   119
         Top             =   780
         Width           =   1560
         Begin ECMixer.MSSlider MSSlider3 
            Height          =   180
            Left            =   75
            TabIndex        =   120
            Top             =   285
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   318
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   90
         Picture         =   "Form5.frx":3F1EC
         ScaleHeight     =   705
         ScaleWidth      =   1560
         TabIndex        =   117
         Top             =   45
         Width           =   1560
         Begin ECMixer.MSSlider MSSlider2 
            Height          =   180
            Left            =   75
            TabIndex        =   118
            Top             =   285
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   318
         End
      End
      Begin ECMixer.MSSlider MSSlider1 
         Height          =   180
         Left            =   180
         TabIndex        =   108
         ToolTipText     =   "manuele fading"
         Top             =   2475
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   318
      End
   End
   Begin VB.PictureBox PicSamplePlayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   6480
      ScaleHeight     =   234
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   106
      Top             =   300
      Width           =   1485
      Begin VB.PictureBox PicEditSampler 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   540
         Picture         =   "Form5.frx":42B76
         ScaleHeight     =   390
         ScaleWidth      =   450
         TabIndex        =   199
         Top             =   1320
         Width           =   450
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -15
         Picture         =   "Form5.frx":43510
         ScaleHeight     =   855
         ScaleWidth      =   1545
         TabIndex        =   187
         Top             =   1785
         Width           =   1545
         Begin ECMixer.MSSlider MSSlider4 
            Height          =   180
            Left            =   45
            TabIndex        =   188
            ToolTipText     =   "sampler volume"
            Top             =   300
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   318
         End
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   132
         Picture         =   "Form5.frx":47ACA
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   135
         Top             =   960
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   504
         Picture         =   "Form5.frx":48350
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   134
         Top             =   960
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   876
         Picture         =   "Form5.frx":48BD6
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   133
         Top             =   960
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   132
         Picture         =   "Form5.frx":4945C
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   132
         Top             =   672
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   504
         Picture         =   "Form5.frx":49CE2
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   131
         Top             =   672
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   876
         Picture         =   "Form5.frx":4A568
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   130
         Top             =   672
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   132
         Picture         =   "Form5.frx":4ADEE
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   129
         Top             =   384
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   504
         Picture         =   "Form5.frx":4B674
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   128
         Top             =   384
         Width           =   450
      End
      Begin VB.PictureBox PicbtnSamplePlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   876
         Picture         =   "Form5.frx":4BEFA
         ScaleHeight     =   345
         ScaleWidth      =   450
         TabIndex        =   127
         Top             =   384
         Width           =   450
      End
      Begin VB.Image Image11 
         Height          =   120
         Left            =   540
         Picture         =   "Form5.frx":4C780
         Top             =   75
         Width           =   270
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "SAMPLE PLAYER"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   144
         Left            =   192
         TabIndex        =   136
         Top             =   216
         Width           =   996
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   15
      Picture         =   "Form5.frx":4C982
      ScaleHeight     =   600
      ScaleWidth      =   6300
      TabIndex        =   89
      Top             =   7035
      Width           =   6300
      Begin VB.PictureBox PicDnPlaylist1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5730
         Picture         =   "Form5.frx":58EA4
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   163
         Top             =   315
         Width           =   285
      End
      Begin VB.PictureBox PicUpPlaylist1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5730
         Picture         =   "Form5.frx":5935A
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   162
         Top             =   30
         Width           =   285
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   16
         Left            =   90
         Picture         =   "Form5.frx":59810
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   103
         ToolTipText     =   "wis lijst"
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   16
         Left            =   75
         Picture         =   "Form5.frx":5A212
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   102
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   15
         Left            =   690
         Picture         =   "Form5.frx":5AC14
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   101
         ToolTipText     =   "laad playlist"
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   15
         Left            =   615
         Picture         =   "Form5.frx":5B616
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   100
         Top             =   945
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   14
         Left            =   1125
         Picture         =   "Form5.frx":5C018
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   99
         Top             =   945
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   13
         Left            =   1650
         Picture         =   "Form5.frx":5CA1A
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   98
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   12
         Left            =   2160
         Picture         =   "Form5.frx":5D41C
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   97
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   11
         Left            =   2685
         Picture         =   "Form5.frx":5DE1E
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   96
         Top             =   975
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   14
         Left            =   1290
         Picture         =   "Form5.frx":5E820
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   95
         ToolTipText     =   "save playlist"
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   13
         Left            =   1875
         Picture         =   "Form5.frx":5F222
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   94
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   12
         Left            =   2475
         Picture         =   "Form5.frx":5FC24
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   93
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   11
         Left            =   3090
         Picture         =   "Form5.frx":60626
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   92
         Top             =   105
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   10
         Left            =   3675
         Picture         =   "Form5.frx":61028
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   91
         Top             =   105
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   10
         Left            =   3195
         Picture         =   "Form5.frx":61A2A
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   90
         Top             =   975
         Width           =   480
      End
      Begin VB.Label LblTotalTimePlaylist1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFE8B6&
         Height          =   210
         Left            =   5295
         TabIndex        =   182
         Top             =   75
         Width           =   315
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   165
         Left            =   5175
         TabIndex        =   104
         ToolTipText     =   "Position of playing track in playlist"
         Top             =   360
         Width           =   405
      End
      Begin VB.Image Image18 
         Height          =   150
         Left            =   4335
         Picture         =   "Form5.frx":6242C
         Top             =   375
         Width           =   165
      End
      Begin VB.Image Image14 
         Height          =   150
         Left            =   4530
         Picture         =   "Form5.frx":625D6
         Top             =   375
         Width           =   165
      End
      Begin VB.Image Image10 
         Height          =   150
         Left            =   4695
         Picture         =   "Form5.frx":62780
         Top             =   375
         Width           =   165
      End
      Begin VB.Image Image9 
         Height          =   150
         Left            =   4890
         Picture         =   "Form5.frx":6292A
         Top             =   375
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   30
      Picture         =   "Form5.frx":62AD4
      ScaleHeight     =   645
      ScaleWidth      =   14385
      TabIndex        =   66
      Top             =   10035
      Width           =   14385
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   25
         Left            =   10350
         Picture         =   "Form5.frx":7E09E
         ScaleHeight     =   390
         ScaleWidth      =   795
         TabIndex        =   195
         Top             =   915
         Width           =   795
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   25
         Left            =   10275
         Picture         =   "Form5.frx":7F120
         ScaleHeight     =   390
         ScaleWidth      =   795
         TabIndex        =   194
         ToolTipText     =   "toon mixer"
         Top             =   105
         Width           =   795
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   24
         Left            =   9390
         Picture         =   "Form5.frx":801A2
         ScaleHeight     =   390
         ScaleWidth      =   795
         TabIndex        =   179
         Top             =   960
         Width           =   795
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   24
         Left            =   9345
         Picture         =   "Form5.frx":81224
         ScaleHeight     =   390
         ScaleWidth      =   795
         TabIndex        =   178
         ToolTipText     =   "toon mixer"
         Top             =   105
         Width           =   795
      End
      Begin VB.PictureBox PicBtnTmp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   75
         Picture         =   "Form5.frx":822A6
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   87
         Top             =   1725
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   630
         Picture         =   "Form5.frx":82CA8
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   86
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   9
         Left            =   4965
         Picture         =   "Form5.frx":836AA
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   85
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   9
         Left            =   5580
         Picture         =   "Form5.frx":840AC
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   84
         ToolTipText     =   "zoek bestand"
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   8
         Left            =   4425
         Picture         =   "Form5.frx":84AAE
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   83
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   8
         Left            =   4980
         Picture         =   "Form5.frx":854B0
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   82
         ToolTipText     =   "toon bestand info"
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   7
         Left            =   4395
         Picture         =   "Form5.frx":85EB2
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   81
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   6
         Left            =   3810
         Picture         =   "Form5.frx":868B4
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   80
         ToolTipText     =   "selectie wissen"
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   3180
         Picture         =   "Form5.frx":872B6
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   79
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   2610
         Picture         =   "Form5.frx":87CB8
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   78
         ToolTipText     =   "alles selecteren"
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   2010
         Picture         =   "Form5.frx":886BA
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   77
         ToolTipText     =   "nummer(s) verwijderen uit lijst"
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   2
         Left            =   1425
         Picture         =   "Form5.frx":890BC
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   76
         ToolTipText     =   "nummer(s) toevoegen aan lijst"
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   7
         Left            =   3885
         Picture         =   "Form5.frx":89ABE
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   75
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   6
         Left            =   3345
         Picture         =   "Form5.frx":8A4C0
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   74
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   2805
         Picture         =   "Form5.frx":8AEC2
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   73
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   2280
         Picture         =   "Form5.frx":8B8C4
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   72
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   1710
         Picture         =   "Form5.frx":8C2C6
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   71
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   2
         Left            =   1170
         Picture         =   "Form5.frx":8CCC8
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   70
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   825
         Picture         =   "Form5.frx":8D6CA
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   69
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox PicbtnDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   0
         Left            =   75
         Picture         =   "Form5.frx":8E0CC
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   68
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox Picbtn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   0
         Left            =   225
         Picture         =   "Form5.frx":8EACE
         ScaleHeight     =   390
         ScaleWidth      =   480
         TabIndex        =   67
         ToolTipText     =   "wis lijst"
         Top             =   120
         Width           =   480
      End
      Begin VB.Label LblTotalTimeGenralList 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000006&
         Caption         =   "0:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFE8B6&
         Height          =   210
         Left            =   13650
         TabIndex        =   184
         Top             =   60
         Width           =   315
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   165
         Left            =   13590
         TabIndex        =   88
         ToolTipText     =   "Position of playing track in playlist"
         Top             =   375
         Width           =   405
      End
      Begin VB.Image LvTrackNext 
         Height          =   150
         Left            =   12930
         Picture         =   "Form5.frx":8F4D0
         Top             =   375
         Width           =   165
      End
      Begin VB.Image LvStopPlay 
         Height          =   150
         Left            =   12735
         Picture         =   "Form5.frx":8F67A
         Top             =   375
         Width           =   165
      End
      Begin VB.Image LvPLayPlay 
         Height          =   150
         Left            =   12570
         Picture         =   "Form5.frx":8F824
         Top             =   375
         Width           =   165
      End
      Begin VB.Image LvTrackPrev 
         Height          =   150
         Left            =   12375
         Picture         =   "Form5.frx":8F9CE
         Top             =   375
         Width           =   165
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2085
      Top             =   1905
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Autoplay"
      Height          =   255
      Left            =   9360
      TabIndex        =   57
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1635
      Top             =   1905
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Autoplay"
      Height          =   255
      Left            =   60
      TabIndex        =   51
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicRightPLayerA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   7935
      Picture         =   "Form5.frx":8FB78
      ScaleHeight     =   234
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   20
      Top             =   300
      Width           =   6495
      Begin VB.Timer Timerdeck2Pause 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2355
         Top             =   1560
      End
      Begin VB.PictureBox PicPlayer2Shuffle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4410
         Picture         =   "Form5.frx":D830A
         ScaleHeight     =   285
         ScaleWidth      =   915
         TabIndex        =   173
         ToolTipText     =   "shuffle mode"
         Top             =   3060
         Width           =   915
      End
      Begin VB.PictureBox PicPLayer2Loop 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3780
         Picture         =   "Form5.frx":D90F4
         ScaleHeight     =   285
         ScaleWidth      =   540
         TabIndex        =   172
         ToolTipText     =   "loop play"
         Top             =   3060
         Width           =   540
      End
      Begin VB.PictureBox PicPLayer2SinglePLay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3135
         Picture         =   "Form5.frx":D993A
         ScaleHeight     =   285
         ScaleWidth      =   540
         TabIndex        =   171
         ToolTipText     =   "single play"
         Top             =   3060
         Width           =   540
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   5505
         Picture         =   "Form5.frx":DA180
         ScaleHeight     =   2400
         ScaleWidth      =   915
         TabIndex        =   124
         Top             =   990
         Width           =   915
         Begin ECMixer.VSlider VSlider2 
            Height          =   2280
            Left            =   270
            TabIndex        =   125
            Top             =   45
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   4022
            MaxValue        =   -100
            Picture         =   "Form5.frx":E14C2
         End
      End
      Begin VB.CommandButton Deck2_Open 
         Caption         =   "Open"
         Height          =   495
         Left            =   2250
         TabIndex        =   116
         Top             =   2205
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.VScrollBar Deck2_Volume 
         Height          =   1575
         Left            =   5250
         Max             =   -100
         TabIndex        =   111
         Top             =   1260
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.CheckBox Deck2_Mute 
         BackColor       =   &H00000000&
         Caption         =   "Mute Player 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   105
         TabIndex        =   65
         Top             =   2790
         Width           =   225
      End
      Begin VB.PictureBox PicPLayerB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   120
         ScaleHeight     =   92
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   319
         TabIndex        =   27
         Top             =   90
         Width           =   4785
         Begin VB.PictureBox PicPlayer2Spectrum 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   60
            ScaleHeight     =   585
            ScaleWidth      =   1095
            TabIndex        =   28
            Top             =   765
            Width           =   1095
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0000C000&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   17
               Left            =   840
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00C0FFC0&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   16
               Left            =   600
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   15
               Left            =   0
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H000080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   14
               Left            =   120
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0080C0FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   13
               Left            =   240
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   12
               Left            =   360
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   11
               Left            =   480
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0080FF80&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   10
               Left            =   720
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00008000&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   9
               Left            =   960
               Top             =   330
               Width           =   90
            End
         End
         Begin ECMixer.MorphDisplay MorphLCDElapsedTimeB 
            Height          =   456
            Left            =   924
            TabIndex        =   29
            Top             =   300
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   794
            BorderWidth     =   0
            NumDigits       =   4
            NumDigitsExp    =   2
            SegmentLitColorNeg=   255
            Value           =   "000000"
            XOffsetExp      =   72
            YOffsetExp      =   10
         End
         Begin ECMixer.MorphDisplay MorphLCDRemainingTimeB 
            Height          =   450
            Left            =   2460
            TabIndex        =   30
            Top             =   300
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
            BorderWidth     =   0
            NumDigits       =   4
            NumDigitsExp    =   2
            SegmentLitColorNeg=   255
            Value           =   "000000"
            XOffsetExp      =   72
            YOffsetExp      =   10
         End
         Begin ECMixer.MorphDisplay MorphLCDTrackB 
            Height          =   450
            Left            =   60
            TabIndex        =   31
            Top             =   315
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   794
            BorderWidth     =   0
            InterDigitGap   =   4
            InterDigitGapExp=   0
            NumDigits       =   3
            NumDigitsExp    =   0
            SegmentHeightExp=   8
            SegmentLitColorNeg=   255
            Value           =   "000"
            XOffsetExp      =   90
            YOffsetExp      =   13
         End
         Begin ECMixer.ECScrollingText ECScrollingText2 
            Height          =   285
            Left            =   1170
            TabIndex        =   197
            Top             =   1035
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   503
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            Index           =   1
            X1              =   68
            X2              =   316
            Y1              =   88
            Y2              =   88
         End
         Begin VB.Label LblPlayerB 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   " PLAYER B "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B5A27B&
            Height          =   150
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Deck2_File 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "<NO FILE>"
            ForeColor       =   &H00C1872F&
            Height          =   480
            Left            =   1230
            OLEDropMode     =   1  'Manual
            TabIndex        =   35
            Top             =   840
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "TRACK"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001086F7&
            Height          =   150
            Left            =   210
            TabIndex        =   34
            Top             =   165
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "REMAINING"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001086F7&
            Height          =   150
            Left            =   2475
            TabIndex        =   33
            Top             =   165
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "ELAPSED"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001086F7&
            Height          =   150
            Left            =   960
            TabIndex        =   32
            Top             =   165
            Width           =   540
         End
      End
      Begin VB.PictureBox PicStopPL2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   2970
         Picture         =   "Form5.frx":E1B68
         ScaleHeight     =   825
         ScaleWidth      =   990
         TabIndex        =   26
         Top             =   2055
         Width           =   990
      End
      Begin VB.PictureBox PicPlayPL2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   4125
         Picture         =   "Form5.frx":E46A2
         ScaleHeight     =   825
         ScaleWidth      =   990
         TabIndex        =   25
         ToolTipText     =   "play-pause"
         Top             =   2055
         Width           =   990
      End
      Begin VB.PictureBox PicLightPlayPL2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   4485
         Picture         =   "Form5.frx":E71DC
         ScaleHeight     =   105
         ScaleWidth      =   270
         TabIndex        =   24
         Top             =   1725
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicLightStopPL2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   3360
         Picture         =   "Form5.frx":E73A6
         ScaleHeight     =   90
         ScaleWidth      =   270
         TabIndex        =   23
         Top             =   1725
         Width           =   270
      End
      Begin VB.PictureBox PicTrackLeftPlayerB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   360
         Picture         =   "Form5.frx":E7538
         ScaleHeight     =   360
         ScaleWidth      =   555
         TabIndex        =   22
         ToolTipText     =   "vorig nummer"
         Top             =   2055
         Width           =   555
      End
      Begin VB.PictureBox PicTrackRightPlayerB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1215
         Picture         =   "Form5.frx":E7FFA
         ScaleHeight     =   360
         ScaleWidth      =   555
         TabIndex        =   21
         ToolTipText     =   "volgend nummer"
         Top             =   2055
         Width           =   555
      End
      Begin ECMixer.PBarY ProgressBar2 
         Height          =   255
         Left            =   5010
         TabIndex        =   140
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         BackColor       =   0
         Style           =   1
      End
      Begin ECMixer.ECSlider ECSlider2 
         Height          =   180
         Left            =   120
         TabIndex        =   174
         Top             =   3135
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   318
         PictureBack     =   "Form5.frx":E8ABC
         PictureProgress =   "Form5.frx":E8F7D
         Bar             =   "Form5.frx":E943E
         BarOver         =   "Form5.frx":E9852
         BarDown         =   "Form5.frx":E9C66
         BackColor       =   0
         Position        =   1
      End
      Begin VB.CheckBox ChkRandomPLayer2 
         Caption         =   "Check3"
         Height          =   210
         Left            =   4935
         TabIndex        =   175
         Top             =   3075
         Width           =   210
      End
      Begin VB.CheckBox ChkLoopPLayer2 
         Caption         =   "Check3"
         Height          =   195
         Left            =   3870
         TabIndex        =   176
         Top             =   3090
         Width           =   180
      End
      Begin VB.CheckBox ChkSinglePLayer2 
         Caption         =   "Check4"
         Height          =   195
         Left            =   3225
         TabIndex        =   177
         Top             =   3105
         Width           =   210
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "VOLUME"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   5610
         TabIndex        =   126
         Top             =   720
         Width           =   525
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   335
         X2              =   335
         Y1              =   22
         Y2              =   30
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   335
         X2              =   407
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   407
         X2              =   407
         Y1              =   30
         Y2              =   22
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   367
         X2              =   367
         Y1              =   30
         Y2              =   22
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "0%       50%     100%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4950
         TabIndex        =   112
         Top             =   450
         Width           =   1470
      End
      Begin VB.Label LblMute 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "MUTE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   150
         Index           =   1
         Left            =   330
         TabIndex        =   64
         Top             =   2835
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "STOP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   150
         Left            =   3315
         TabIndex        =   39
         Top             =   1830
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "PLAY PAUSE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   150
         Left            =   4215
         TabIndex        =   38
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "TRACK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   150
         Left            =   930
         TabIndex        =   37
         Top             =   1830
         Width           =   420
      End
   End
   Begin VB.PictureBox PicLeftPLayerA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   15
      Picture         =   "Form5.frx":EA07A
      ScaleHeight     =   234
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   0
      Top             =   300
      Width           =   6480
      Begin VB.Timer Timerdeck1Pause 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2520
         Top             =   1605
      End
      Begin ECMixer.ECSlider ECSlider1 
         Height          =   180
         Left            =   120
         TabIndex        =   161
         Top             =   3135
         Width           =   2835
         _ExtentX        =   318
         _ExtentY        =   5001
         PictureBack     =   "Form5.frx":13280C
         PictureProgress =   "Form5.frx":132CCD
         Bar             =   "Form5.frx":13318E
         BarOver         =   "Form5.frx":1335A2
         BarDown         =   "Form5.frx":1339B6
         BackColor       =   0
         Position        =   1
      End
      Begin ECMixer.PBarY ProgressBar1 
         Height          =   255
         Left            =   5115
         TabIndex        =   139
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   450
         BackColor       =   0
         Style           =   1
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2352
         Left            =   5505
         Picture         =   "Form5.frx":133DCA
         ScaleHeight     =   2355
         ScaleWidth      =   915
         TabIndex        =   121
         Top             =   990
         Width           =   912
         Begin ECMixer.VSlider VSlider1 
            Height          =   2280
            Left            =   270
            TabIndex        =   122
            Top             =   45
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   4022
            MaxValue        =   -100
            Picture         =   "Form5.frx":13B10C
         End
      End
      Begin VB.VScrollBar Deck1_Volume 
         Height          =   1500
         Left            =   5280
         Max             =   -100
         TabIndex        =   110
         Top             =   1380
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.CommandButton Deck1_Open 
         Caption         =   "Open"
         Height          =   495
         Left            =   2268
         TabIndex        =   56
         Top             =   2112
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox Deck1_Mute 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "Mute"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   105
         MaskColor       =   &H80000006&
         TabIndex        =   50
         Top             =   2790
         Width           =   225
      End
      Begin VB.PictureBox PicTrackRightPlayerA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1215
         Picture         =   "Form5.frx":13B7B2
         ScaleHeight     =   360
         ScaleWidth      =   555
         TabIndex        =   17
         ToolTipText     =   "volgend nummer"
         Top             =   2055
         Width           =   555
      End
      Begin VB.PictureBox PicTrackLeftPlayerA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   360
         Picture         =   "Form5.frx":13C274
         ScaleHeight     =   360
         ScaleWidth      =   555
         TabIndex        =   16
         ToolTipText     =   "vorig nummer"
         Top             =   2055
         Width           =   555
      End
      Begin VB.PictureBox PicLightStopPL1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   3360
         Picture         =   "Form5.frx":13CD36
         ScaleHeight     =   90
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   1725
         Width           =   270
      End
      Begin VB.PictureBox PicLightPlayPL1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   4476
         Picture         =   "Form5.frx":13CEC8
         ScaleHeight     =   105
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   1725
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox PicPlayPL1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   4125
         Picture         =   "Form5.frx":13D092
         ScaleHeight     =   825
         ScaleWidth      =   990
         TabIndex        =   11
         ToolTipText     =   "play-pause"
         Top             =   2055
         Width           =   990
      End
      Begin VB.PictureBox PicStopPL1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   2970
         Picture         =   "Form5.frx":13FBCC
         ScaleHeight     =   825
         ScaleWidth      =   990
         TabIndex        =   10
         Top             =   2055
         Width           =   990
      End
      Begin VB.PictureBox PicPLayerA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1380
         Left            =   120
         ScaleHeight     =   92
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   319
         TabIndex        =   1
         Top             =   60
         Width           =   4785
         Begin VB.PictureBox PicPlayer1Spectrum 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   45
            ScaleHeight     =   585
            ScaleWidth      =   1095
            TabIndex        =   19
            Top             =   765
            Width           =   1095
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00008000&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   8
               Left            =   960
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0080FF80&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   6
               Left            =   720
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   4
               Left            =   480
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   3
               Left            =   360
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0080C0FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   2
               Left            =   240
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H000080FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   1
               Left            =   120
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   0
               Left            =   0
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00C0FFC0&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   5
               Left            =   600
               Top             =   330
               Width           =   90
            End
            Begin VB.Shape sp 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H0000C000&
               FillStyle       =   0  'Solid
               Height          =   255
               Index           =   7
               Left            =   840
               Top             =   330
               Width           =   90
            End
         End
         Begin ECMixer.MorphDisplay MorphLCDElapsedTimeA 
            Height          =   450
            Left            =   945
            TabIndex        =   2
            Top             =   300
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
            BorderWidth     =   0
            NumDigits       =   4
            NumDigitsExp    =   2
            SegmentLitColorNeg=   255
            Value           =   "000000"
            XOffsetExp      =   72
            YOffsetExp      =   10
         End
         Begin ECMixer.MorphDisplay MorphLCDRemainingTimeA 
            Height          =   450
            Left            =   2460
            TabIndex        =   3
            Top             =   300
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
            BorderWidth     =   0
            NumDigits       =   4
            NumDigitsExp    =   2
            SegmentLitColorNeg=   255
            Value           =   "000000"
            XOffsetExp      =   72
            YOffsetExp      =   10
         End
         Begin ECMixer.MorphDisplay MorphLCDTrackA 
            Height          =   450
            Left            =   45
            TabIndex        =   4
            Top             =   315
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   794
            BorderWidth     =   0
            InterDigitGap   =   4
            InterDigitGapExp=   0
            NumDigits       =   3
            NumDigitsExp    =   0
            SegmentHeightExp=   8
            SegmentLitColorNeg=   255
            Value           =   "000"
            XOffsetExp      =   90
            YOffsetExp      =   13
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            Index           =   0
            X1              =   68
            X2              =   316
            Y1              =   88
            Y2              =   88
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "ELAPSED"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001086F7&
            Height          =   150
            Left            =   960
            TabIndex        =   9
            Top             =   165
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "REMAINING"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001086F7&
            Height          =   156
            Left            =   2460
            TabIndex        =   8
            Top             =   168
            Width           =   708
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "TRACK"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001086F7&
            Height          =   150
            Left            =   210
            TabIndex        =   7
            Top             =   165
            Width           =   420
         End
         Begin VB.Label Deck1_File 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "<NO FILE>"
            ForeColor       =   &H00C1872F&
            Height          =   195
            Left            =   1275
            OLEDropMode     =   1  'Manual
            TabIndex        =   6
            Top             =   825
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VB.Label LblPlayerA 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   " PLAYER A "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B5A27B&
            Height          =   150
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.PictureBox PicPLayer1SinglePLay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3135
         Picture         =   "Form5.frx":142706
         ScaleHeight     =   285
         ScaleWidth      =   540
         TabIndex        =   166
         ToolTipText     =   "single play"
         Top             =   3090
         Width           =   540
      End
      Begin VB.PictureBox PicPlayer1Shuffle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4410
         Picture         =   "Form5.frx":142F4C
         ScaleHeight     =   285
         ScaleWidth      =   915
         TabIndex        =   164
         ToolTipText     =   "shuffle mode"
         Top             =   3090
         Width           =   915
      End
      Begin VB.CheckBox ChkSinglePLayer1 
         Caption         =   "Check4"
         Height          =   195
         Left            =   3210
         TabIndex        =   170
         Top             =   3165
         Width           =   210
      End
      Begin VB.CheckBox ChkRandomPLayer1 
         Caption         =   "Check3"
         Height          =   210
         Left            =   4920
         TabIndex        =   160
         Top             =   3135
         Width           =   210
      End
      Begin VB.PictureBox PicPLayer1Loop 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3780
         Picture         =   "Form5.frx":143D36
         ScaleHeight     =   285
         ScaleWidth      =   540
         TabIndex        =   165
         ToolTipText     =   "loop play"
         Top             =   3090
         Width           =   540
      End
      Begin VB.CheckBox ChkLoopPLayer1 
         Caption         =   "Check3"
         Height          =   195
         Left            =   3855
         TabIndex        =   169
         Top             =   3150
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "VOLUME"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   150
         Left            =   5685
         TabIndex        =   123
         Top             =   720
         Width           =   525
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   341
         X2              =   341
         Y1              =   24
         Y2              =   32
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   342
         X2              =   414
         Y1              =   32
         Y2              =   32
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   414
         X2              =   414
         Y1              =   31
         Y2              =   23
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   373
         X2              =   373
         Y1              =   32
         Y2              =   24
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "0%       50%     100%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4980
         TabIndex        =   109
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label LblMute 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "MUTE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   150
         Index           =   0
         Left            =   315
         TabIndex        =   63
         Top             =   2835
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "TRACK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   150
         Left            =   930
         TabIndex        =   18
         Top             =   1830
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "PLAY PAUSE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   150
         Left            =   4215
         TabIndex        =   15
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "STOP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   150
         Left            =   3315
         TabIndex        =   14
         Top             =   1830
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2400
      Left            =   3060
      TabIndex        =   138
      Top             =   7650
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16771254
      BackColor       =   -2147483642
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   926
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Artist"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Album"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Genre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Track No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tracks Total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Year"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Duration"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Bit Rate"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Comments"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   6585
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   186
      Top             =   5235
      Width           =   1215
   End
   Begin MSComctlLib.ListView LvSearch 
      Height          =   2400
      Left            =   3060
      TabIndex        =   192
      Top             =   7635
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16771254
      BackColor       =   -2147483642
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   926
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Artist"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Album"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Genre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Track No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tracks Total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Year"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Duration"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Bit Rate"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Comments"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   3540
      Left            =   165
      TabIndex        =   40
      Top             =   3720
      Visible         =   0   'False
      Width           =   4050
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   255
         Left            =   3696
         TabIndex        =   60
         Top             =   3084
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   255
         Left            =   96
         TabIndex        =   59
         Top             =   3084
         Width           =   255
      End
      Begin VB.HScrollBar Cross_Fader 
         Height          =   240
         Left            =   525
         Max             =   100
         Min             =   -100
         TabIndex        =   47
         Top             =   2490
         Value           =   100
         Width           =   2760
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "off"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1704
         TabIndex        =   46
         Top             =   1104
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "on"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1704
         TabIndex        =   45
         Top             =   876
         Width           =   495
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   2400
         Top             =   2160
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   1560
         Top             =   2160
      End
      Begin VB.Timer timerfade2 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   2040
         Top             =   2160
      End
      Begin VB.Timer timerfade1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1080
         Top             =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "> X"
         Height          =   285
         Left            =   3336
         TabIndex        =   44
         Top             =   2460
         Width           =   396
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X <"
         Height          =   270
         Left            =   192
         TabIndex        =   43
         Top             =   2475
         Width           =   360
      End
      Begin VB.PictureBox Slider2 
         Height          =   276
         Left            =   2184
         ScaleHeight     =   210
         ScaleWidth      =   1050
         TabIndex        =   41
         Top             =   960
         Width           =   1104
      End
      Begin VB.PictureBox Slider1 
         Height          =   276
         Left            =   504
         ScaleHeight     =   210
         ScaleWidth      =   1050
         TabIndex        =   58
         Top             =   948
         Width           =   1104
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "L   Balance  R"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   2
         Left            =   2268
         TabIndex        =   42
         Top             =   1440
         Width           =   984
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "L   Balance  R  Mono"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   0
         Left            =   516
         TabIndex        =   49
         Top             =   1452
         Width           =   1464
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Crossfader"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   48
         Top             =   2250
         Width           =   3690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   1950
         X2              =   1950
         Y1              =   2850
         Y2              =   2310
      End
   End
   Begin MSComctlLib.ListView LvPlaylist1 
      Height          =   3195
      Left            =   60
      TabIndex        =   142
      Top             =   3810
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16771254
      BackColor       =   -2147483642
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Artist"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Album"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Genre"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Track No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Tracks Total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Year"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Duration"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Bit Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Comments"
         Object.Width           =   0
      EndProperty
   End
   Begin WMPLibCtl.WindowsMediaPlayer LVmediaPLayer 
      Height          =   405
      Left            =   8490
      TabIndex        =   181
      Top             =   8085
      Width           =   1560
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2752
      _cy             =   714
   End
   Begin WMPLibCtl.WindowsMediaPlayer Deck1 
      Height          =   450
      Left            =   75
      TabIndex        =   159
      Top             =   4185
      Width           =   2340
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4128
      _cy             =   794
   End
   Begin WMPLibCtl.WindowsMediaPlayer SampPlayer 
      Height          =   450
      Left            =   3195
      TabIndex        =   105
      Top             =   4095
      Width           =   2340
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4128
      _cy             =   794
   End
   Begin VB.Label Deck2_Time 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1872F&
      Height          =   120
      Left            =   10710
      TabIndex        =   62
      Top             =   3855
      Width           =   720
   End
   Begin VB.Label Deck2_Remain 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1872F&
      Height          =   120
      Left            =   10710
      TabIndex        =   61
      Top             =   3975
      Width           =   720
   End
   Begin WMPLibCtl.WindowsMediaPlayer Deck2 
      Height          =   450
      Left            =   9375
      TabIndex        =   55
      Top             =   4110
      Width           =   2340
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4128
      _cy             =   794
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "lbl"
      Height          =   195
      Left            =   2595
      TabIndex        =   54
      Top             =   3840
      Width           =   150
   End
   Begin VB.Label Deck1_Remain 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1872F&
      Height          =   210
      Left            =   1440
      TabIndex        =   53
      Top             =   3975
      Width           =   630
   End
   Begin VB.Label Deck1_Time 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1872F&
      Height          =   255
      Left            =   1440
      TabIndex        =   52
      Top             =   3810
      Width           =   630
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

DefLng A-Z

Private DragLV                      As ListItem 'The item being dragged


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Const CF_BITMAP As Long = 2
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

Private Const S_OTHER As String = "Other"

Private Const FILTER_BMP As String = "*.bmp;*.dib"
Private Const FILTER_GIF As String = "*.gif"
Private Const FILTER_JPEG As String = "*.jpeg;*.jpg;*.jpe;*.jfif;*.jfi;*.jif"
Private Const FILTER_PNG As String = "*.png"
Private Const FILTER_SUPPORTED As String = FILTER_BMP & ";" & FILTER_GIF & ";" & FILTER_JPEG & ";" & FILTER_PNG

Private Const MNU_COPY As Long = 0
Private Const MNU_PASTE As Long = 1

Private Const PASTE_TXT_1 As String = "&Paste"
Private Const PASTE_TXT_2 As String = PASTE_TXT_1 & " (this will change the current image)"

Dim myWindowState As Integer
Dim bInitialized As Boolean


Dim Pos As Single
Dim CutS As String, CutI As Long 'Current song string
Dim BPos As Integer
Dim PlayerAStopped As Boolean
Dim PlayerBStopped As Boolean

Dim vClrs(1 To 9) As Long
Dim P_Val&, m_Max&
Dim LstPeak&, WtPeak$
Private BPMArray(14) As Single
Private LastBPM
Private MaxBPMs
Private BPM As Integer
Private OldBPM As Integer
Dim volume As Integer
Dim buffaddress As Long
Dim audbytearray As AUDINPUTARRAY
Dim retVal As Integer
Private OutBuffer As String
Dim Deck1Pause As Boolean
Dim Deck2Pause As Boolean

Function XTwips(Pixels) As Integer

    'convert pixels to twips
    XTwips = Pixels * Screen.TwipsPerPixelX

End Function

Function YTwips(Pixels) As Integer

    'convert pixels to twips
    YTwips = Pixels * Screen.TwipsPerPixelY

End Function

Private Sub bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Choosing = True
        OrigX = X
        OrigY = Y
    End If
    If Button = 2 Then
        'Extras.PopupMenu Extras.Options
    End If
End Sub

Private Sub bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XCaseVar As Integer, YCaseVar As Integer
    If Button = 1 And Choosing = True Then
        XCaseVar = 1
        YCaseVar = 1
        If Me.Top + YTwips(Y - OrigY) > TopVar - YTwips(10) - Me.Height And Me.Top + YTwips(Y - OrigY2) < TopVar + YTwips(10) Then

        If Me.Left + XTwips(X - OrigX) + Me.Width > LeftVar - XTwips(10) And Me.Left + XTwips(X - OrigX) + Me.Width < LeftVar + XTwips(10) Then
            XCaseVar = 3
        End If
        If Me.Left + XTwips(X - OrigX) > LeftVar - XTwips(10) And Me.Left + XTwips(X - OrigX) < LeftVar + XTwips(10) Then
            XCaseVar = 4
        End If
        End If
        If Me.Left + XTwips(X - OrigX) > LeftVar - XTwips(10) - Me.Width And Me.Left + XTwips(X - OrigX2) < LeftVar + XTwips(10) Then
        If Me.Top + YTwips(Y - OrigY) + Me.Height > TopVar - YTwips(10) And Me.Top + YTwips(Y - OrigY) + Me.Height < TopVar + YTwips(10) Then
            YCaseVar = 3
        End If
        If Me.Top + YTwips(Y - OrigY) > TopVar - YTwips(10) And Me.Top + YTwips(Y - OrigY) < TopVar + YTwips(10) Then
            YCaseVar = 4
        End If
        End If
        End If
        
        XSnapped = True
        YSnapped = True
        Select Case XCaseVar
            Case 1
                Me.Left = Me.Left + XTwips(X - OrigX)
            Case 2
                Me.Left = LeftVar
            Case 3
                Me.Left = LeftVar - Me.Width
            Case 4
                Me.Left = LeftVar
        End Select
        Select Case YCaseVar
            Case 1
                Me.Top = Me.Top + YTwips(Y - OrigY)

            Case 2
                Me.Top = TopVar
            Case 3
                Me.Top = TopVar - Me.Height
            Case 4
                Me.Top = TopVar
        End Select
        Me.Refresh
        


End Sub



Private Function CalcTotalTime(lv As ListView) As String
Dim i As Long
Dim txtSec As Long
Dim Hours, Minutes, Seconds As Long
Dim X As String
On Error Resume Next
txtSec = 0
       For i = 1 To lv.ListItems.Count
        
        X = lv.ListItems(i).SubItems(9)
        If X <> "" Then
          If Len(X) <= 4 Then X = "00:0" & X
            
          Hours = Mid(X, 1, 2) * 3600
          Minutes = (Mid(X, 4, 2) * 60)
          Seconds = Right(X, 2)
      
          txtSec = txtSec + (Hours + Minutes + Seconds)
        End If
      Next i


    CalcTotalTime = SecondsToText(txtSec)
    On Error GoTo 0
End Function

Function SecondsToText(Seconds) As String
Dim bAddComma As Boolean
Dim Result As String
Dim sTemp As String

If Seconds <= 0 Or Not IsNumeric(Seconds) Then
     SecondsToText = "0:00"
     Exit Function
End If

Seconds = Fix(Seconds)

If Seconds >= 86400 Then
  days = Fix(Seconds / 86400)
Else
  days = 0
End If

If Seconds - (days * 86400) >= 3600 Then
  Hours = Fix((Seconds - (days * 86400)) / 3600)
Else
  Hours = 0
End If

If Seconds - (Hours * 3600) - (days * 86400) >= 60 Then
 Minutes = Fix((Seconds - (Hours * 3600) - (days * 86400)) / 60)
Else
 Minutes = 0
End If

Seconds = Seconds - (Minutes * 60) - (Hours * 3600) - _
   (days * 86400)

If Seconds > 0 Then Result = Seconds

If Minutes > 0 Then
    sTemp = Minutes & ":"
    Result = sTemp & Result
End If

If Hours > 0 Then
    sTemp = Hours & ":"
    Result = sTemp & Result
End If

If days > 0 Then
    sTemp = days & ", "
    Result = sTemp & Result
End If
Debug.Print Result
SecondsToText = Result
End Function


Function AutoS(Number)
    'If Number = 1 Then AutoS = "" Else AutoS = "s"
End Function

Private Sub bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Choosing = False
    End If
End Sub

Private Sub Check1_Click()
Deck1.settings.autoStart = Not Deck1.settings.autoStart
End Sub

Private Sub Check2_Click()
Deck2.settings.autoStart = Not Deck2.settings.autoStart
End Sub

Private Sub ChkAutomaticMixing_Click()
PicResetFader.Enabled = False
Timer3.Enabled = Not Timer3.Enabled
Command3.Enabled = Not Command3.Enabled
Command2.Enabled = Not Command2.Enabled
End Sub

Private Sub Cmdtest_Click()
  SampPlayer.URL = "C:\MyDocuments\Help programming\Mixer with crossfader - V3.50 Remixed By Stuntmaster (updated)\wav\PECHE A.wav"
  SampPlayer.Controls.play

End Sub

Private Sub Command8_Click()
  For i = 0 To 14
        BPMArray(i) = 0
  Next
  MaxBPMs = 0
  Label8.Caption = format(OldBPM / 100, "##0.00")
  'Command7.SetFocus
End Sub

Private Sub Cross_Fader_Change()
  ' Right, this is where the crossfading is done, 2 lines of code! Simple!
  If Cross_Fader.Value > 0 Then Deck1_Volume.Value = (100 - Cross_Fader.Value) - 100
  If Cross_Fader.Value < 0 Then Deck2_Volume.Value = Cross_Fader.Value
  MSSlider1.Value = Cross_Fader.Value
End Sub

Private Sub Cross_Fader_Scroll()
  Cross_Fader_Change
  timerfade1.Enabled = False
  timerfade2.Enabled = False
End Sub



Private Sub Deck1_File_Change()
  On Error Resume Next
  ECScrollingText1.Text = UCase(LvPlaylist1.ListItems(LvPlaylist1.selectedItem.Index).SubItems(3) & " - " & LvPlaylist1.ListItems(LvPlaylist1.selectedItem.Index).SubItems(2))
  Call ECScrollingText1.ToonText(True)

End Sub

Private Sub Deck1_Mute_Click()
  If Deck1_Mute.Value = 1 Then Deck1.settings.mute = True
  If Deck1_Mute.Value = 0 Then Deck1.settings.mute = False
End Sub


Private Sub Deck1_Open_Click()
  On Error GoTo Error
  Dialog.CancelError = True 'This is to stop the track resetting when playing
                            'if cancel is pressed
  Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
  Dialog.ShowOpen
  
  Deck1.URL = Dialog.filename
  'Visual basic 6 users may want to get rid of the module...since it is a feature
  'that is already on VB6 (InStrRev)
  Deck1_File.Caption = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
  If Dialog.filename = "" Then Deck1_File.Caption = "<NO FILE>"
  Exit Sub
  
Error:
  If err.Number <> 32755 Then
    MsgBox "Error loading file - " & err.Number & " : " & err.Description
  Else
  End If
End Sub


Private Sub Deck1_PlayStateChange(ByVal NewState As Long)
  Dim xx As Integer
  Dim xlast As Integer
  
 
  
  'If NewState = 0 Then Label2.Caption = "Idle": Shape1.Visible = False
  If NewState = 8 Then ' "stopped"
    'Timer1.Enabled = False
    PicPlayer1Spectrum.Visible = False

  End If

  If NewState = 3 Then
      Timerdeck1Pause.Enabled = False: PicPlayer1Spectrum.Visible = True: 'Timer1.Enabled = True  ' "Playing"
      ECScrollingText1.Scrolling = True
      PicLightPlayPL1.Visible = True
      PicLightStopPL1.Visible = False

  End If
  If NewState = 2 Then Timerdeck1Pause.Enabled = True: PicPlayer1Spectrum.Visible = False '"Paused"
  'If NewState = 3 Then Label2.Caption = "Waiting..": Shape1.Visible = False
  'If NewState = 4 Then Label2.Caption = "Scan >>": Shape1.Visible = True
  'If NewState = 5 Then Label2.Caption = "<< Scan": Shape1.Visible = True
  'If NewState = 6 Then Label2.Caption = "Idle": Shape1.Visible = False
End Sub

Private Sub Deck1_Volume_Change()
  Deck1.settings.volume = Deck1_Volume.Value + 100
  MSSlider2.Value = Deck1_Volume.Value
  VSlider1.Value = Deck1_Volume.Value
End Sub

Private Sub Deck1_Volume_Scroll()
  Deck1_Volume_Change
End Sub

Private Sub Deck2_File_Change()
      ECScrollingText2.Text = UCase(LvPlaylist2.ListItems(LvPlaylist2.selectedItem.Index).SubItems(3) & " - " & LvPlaylist2.ListItems(LvPlaylist2.selectedItem.Index).SubItems(2))
      Call ECScrollingText2.ToonText(True)
End Sub

Private Sub Deck2_Mute_Click()
  If Deck2_Mute.Value = 1 Then Deck2.settings.mute = True
  If Deck2_Mute.Value = 0 Then Deck2.settings.mute = False
End Sub

Private Sub Deck2_Open_Click()
  On Error GoTo Error
  Dialog.CancelError = True
  Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
  Dialog.ShowOpen
  
  Deck2.URL = Dialog.filename
  'Visual basic 6 users may want to get rid of the module...since it is a feature
  'that is already on VB6 (InStrRev)
  Deck2_File.Caption = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
  If Dialog.filename = "" Then Deck2_File.Caption = "<NO FILE>"
  Exit Sub
  
Error:
  If err.Number <> 32755 Then ' Cancel was pressed?
    MsgBox "Error loading file - " & err.Number & " : " & err.Description
  Else
  End If
End Sub

Private Sub Deck2_PlayStateChange(ByVal NewState As Long)
  Dim xx As Integer
  Dim xlast As Integer
  
  'If NewState = 0 Then Label2.Caption = "Idle": Shape1.Visible = False
  If NewState = 8 Then ' "stopped"
    'Timer1.Enabled = False:
    PicPlayer2Spectrum.Visible = False
  
    If PlayerBStopped = False Then

    End If
  End If
  
  If NewState = 3 Then
    Timerdeck2Pause.Enabled = False: PicPlayer2Spectrum.Visible = True: 'Timer1.Enabled = True  ' "Playing"
    ECScrollingText2.Scrolling = True
    PicLightPlayPL2.Visible = True
    PicLightStopPL2.Visible = False

  End If
  If NewState = 2 Then Timerdeck2Pause.Enabled = True: PicPlayer2Spectrum.Visible = False '"Paused"
  'If NewState = 3 Then Label2.Caption = "Waiting..": Shape1.Visible = False
  'If NewState = 4 Then Label2.Caption = "Scan >>": Shape1.Visible = True
  'If NewState = 5 Then Label2.Caption = "<< Scan": Shape1.Visible = True
  'If NewState = 6 Then Label2.Caption = "Idle": Shape1.Visible = False

End Sub

Private Sub Deck2_Volume_Change()
  Deck2.settings.volume = Deck2_Volume.Value + 100
  MSSlider3.Value = Deck2_Volume.Value
  VSlider2.Value = Deck2_Volume.Value
  Label2.Caption = Deck2_Volume.Value + 100
End Sub

Private Sub Deck2_Volume_Scroll()
  Deck2_Volume_Change
End Sub

Private Sub ECSlider1_Change(Value As Long)
  Deck1.Controls.currentPosition = ECSlider1.Value
End Sub

Private Sub ECSlider2_Change(Value As Long)
  Deck2.Controls.currentPosition = ECSlider2.Value
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Dim i As Long
    Dim t As Long
    Dim strT As String
    
    

    
      Me.Line (2, 2)-(Me.ScaleWidth - 4, 2), &H5A5952, B
  Me.Line (2, 2)-(2, Me.ScaleHeight - 4), &H5A5952, B
  Me.Line (Me.ScaleWidth - 4, 2)-(Me.ScaleWidth - 4, Me.ScaleHeight - 2), &H5A5952, B
  Me.Line (2, Me.ScaleHeight - 2)-(Me.ScaleWidth - 4, Me.ScaleHeight - 2), &H5A5952, B



    vClrs(1) = vbGreen
    vClrs(2) = vbGreen
    vClrs(3) = vbGreen
    vClrs(4) = vbGreen
    vClrs(5) = vbYellow
    vClrs(6) = vbYellow
    vClrs(7) = vbYellow
    vClrs(8) = vbRed
    vClrs(9) = vbRed

    P_Val = 0
    m_Max = 100
    
    MSSlider1.min = -100
    MSSlider1.Max = 100
    MSSlider1.LowColor = &H846563
    MSSlider1.HiColor = &H846563
    MSSlider1.ColorShift = True
    MSSlider1.ColInit
    MSSlider1.Value = 0
  
    MSSlider2.min = -100
    MSSlider2.Max = 0
    MSSlider2.LowColor = &H846563
    MSSlider2.HiColor = &H846563
    MSSlider2.ColorShift = True
    MSSlider2.ColInit
    MSSlider2.Value = 0
    
    MSSlider3.min = -100
    MSSlider3.Max = 0
    MSSlider3.LowColor = &H846563
    MSSlider3.HiColor = &H846563
    MSSlider3.ColorShift = True
    MSSlider3.ColInit
    MSSlider3.Value = 0
    
    
    
  MSSlider4.Max = 100
  MSSlider4.Value = SampPlayer.settings.volume
  MSSlider4.LowColor = &H846563
  MSSlider4.HiColor = &H846563
  MSSlider4.ColorShift = True
  MSSlider4.ColInit
  
  PicSamplePlayer.Line (2, 2)-(PicSamplePlayer.ScaleWidth - 2, 2), &H5A5952, B
  PicSamplePlayer.Line (2, 2)-(2, PicSamplePlayer.ScaleHeight - 2), &H5A5952, B
  PicSamplePlayer.Line (PicSamplePlayer.ScaleWidth - 2, 2)-(PicSamplePlayer.ScaleWidth - 2, PicSamplePlayer.ScaleHeight - 2), &H5A5952, B
  PicSamplePlayer.Line (2, PicSamplePlayer.ScaleHeight - 2)-(PicSamplePlayer.ScaleWidth - 2, PicSamplePlayer.ScaleHeight - 2), &H5A5952, B
  'Draw frame time slider
  PicLeftPLayerA.Line (6, 205)-(202, 205), &H5A5952, B
  PicLeftPLayerA.Line (6, 205)-(6, 225), &H5A5952, B
  PicLeftPLayerA.Line (202, 205)-(202, 225), &H5A5952, B
  PicLeftPLayerA.Line (6, 225)-(202, 225), &H5A5952, B
  'Draw frame time slider
  PicRightPLayerA.Line (6, 205)-(202, 205), &H5A5952, B
  PicRightPLayerA.Line (6, 205)-(6, 225), &H5A5952, B
  PicRightPLayerA.Line (202, 205)-(202, 225), &H5A5952, B
  PicRightPLayerA.Line (6, 225)-(202, 225), &H5A5952, B
  
  
  
  PicLeftPLayerA.Line (2, 2)-(PicLeftPLayerA.ScaleWidth - 2, 2), &H5A5952, B
  PicLeftPLayerA.Line (2, 2)-(2, PicLeftPLayerA.ScaleHeight - 2), &H5A5952, B
  PicLeftPLayerA.Line (PicLeftPLayerA.ScaleWidth - 2, 2)-(PicLeftPLayerA.ScaleWidth - 2, PicLeftPLayerA.ScaleHeight - 2), &H5A5952, B
  PicLeftPLayerA.Line (2, PicLeftPLayerA.ScaleHeight - 2)-(PicLeftPLayerA.ScaleWidth - 2, PicLeftPLayerA.ScaleHeight - 2), &H5A5952, B
  
  PicPLayerA.Line (2, 2)-(PicPLayerA.ScaleWidth - 2, 2), &H5A5952, B
  PicPLayerA.Line (6, 5)-(PicPLayerA.ScaleWidth - 6, 5), &HB5A27B, B
  PicPLayerA.Line (2, 2)-(2, PicPLayerA.ScaleHeight - 2), &H5A5952, B
  PicPLayerA.Line (PicPLayerA.ScaleWidth - 2, 2)-(PicPLayerA.ScaleWidth - 2, PicPLayerA.ScaleHeight - 2), &H5A5952, B
  PicPLayerA.Line (2, PicPLayerA.ScaleHeight - 2)-(PicPLayerA.ScaleWidth - 2, PicPLayerA.ScaleHeight - 2), &H5A5952, B
  LblPlayerA.Left = (PicPLayerA.ScaleWidth - LblPlayerA.Width) / 2
  
  PicRightPLayerA.Line (2, 2)-(PicRightPLayerA.ScaleWidth - 2, 2), &H5A5952, B
  PicRightPLayerA.Line (2, 2)-(2, PicRightPLayerA.ScaleHeight - 2), &H5A5952, B
  PicRightPLayerA.Line (PicRightPLayerA.ScaleWidth - 2, 2)-(PicRightPLayerA.ScaleWidth - 2, PicRightPLayerA.ScaleHeight - 2), &H5A5952, B
  PicRightPLayerA.Line (2, PicRightPLayerA.ScaleHeight - 2)-(PicRightPLayerA.ScaleWidth - 2, PicRightPLayerA.ScaleHeight - 2), &H5A5952, B
  
  PicPLayerB.Line (2, 2)-(PicPLayerB.ScaleWidth - 2, 2), &H5A5952, B
  PicPLayerB.Line (6, 5)-(PicPLayerB.ScaleWidth - 6, 5), &HB5A27B, B
  PicPLayerB.Line (2, 2)-(2, PicPLayerB.ScaleHeight - 2), &H5A5952, B
  PicPLayerB.Line (PicPLayerB.ScaleWidth - 2, 2)-(PicPLayerB.ScaleWidth - 2, PicPLayerB.ScaleHeight - 2), &H5A5952, B
  PicPLayerB.Line (2, PicPLayerB.ScaleHeight - 2)-(PicPLayerB.ScaleWidth - 2, PicPLayerB.ScaleHeight - 2), &H5A5952, B
  LblPlayerB.Left = (PicPLayerB.ScaleWidth - LblPlayerB.Width) / 2
  
  
  bInitialized = True
    
  With ListView1
    .ColumnHeaderIcons = ImageList1
    
    .SortKey = GetSetting(App.Title, "Columns", "SortKey", 0)
    If err Then
        err.Clear
        .SortKey = 0
    End If
    
    .SortOrder = GetSetting(App.Title, "Columns", "SortOrder", lvwAscending)
    If err Then
        err.Clear
        .SortOrder = lvwAscending
    End If
    Resort = False
    GoTo skip
    ShowListViewColumnHeaderSortIcon ListView1
    
    For i = 1 To .ColumnHeaders.Count
        t = GetSetting(App.Title, "ColumnPos", format$(i, "00"), CStr(i))
        If err Then
            err.Clear
            t = i
        End If
        
        .ColumnHeaders(t).Position = CInt(i)
        If err Then err.Clear
        
        .ColumnHeaders(i).Width = GetSetting(Caption, "Columns", format$(i, "00"), .ColumnHeaders(i).Width)
        If Not err Then
            If .ColumnHeaders(i).Width < 200 Then
                .ColumnHeaders(i).Width = 200
            End If
        Else
            err.Clear
        End If
    Next
  End With
  
skip:
  MorphLCDTrackB.Value = "00"
  MorphLCDTrackA.Value = "00"
  MorphLCDElapsedTimeA.Value = "---"
  MorphLCDElapsedTimeB.Value = "---"
  MorphLCDRemainingTimeA.Value = ""
  MorphLCDRemainingTimeB.Value = ""
  
  Cross_Fader.Value = 0
  Me.Caption = "EC MP3 Studio " & App.Major & "." & App.Minor & App.Revision & " By Erwin Christiaens"
  Deck1.settings.autoStart = False
  Deck2.settings.autoStart = False
  SoundMeter.BUFFER_SIZE = 800
  SoundMeter.StartInput
  Deck1Pause = True
  Deck2Pause = True
  Me.Top = 0
  
  
  strT = GetOpstartDir
  Text1 = strT
  Deck1_Volume.Value = Deck1.settings.volume
  Deck2_Volume.Value = Deck2.settings.volume
  Me.Show
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'unload any other form
   
    Dim i As Long
    
    'SaveSetting Caption, "Window", "XSize", XSize
    'SaveSetting Caption, "Window", "YSize", YSize
    'SaveSetting Caption, "Window", "XPos", XPos
    'SaveSetting Caption, "Window", "YPos", YPos
    'SaveSetting Caption, "Window", "State", myWindowState
    
    With ListView1
        SaveSetting App.Title, "Columns", "SortKey", .SortKey
        SaveSetting App.Title, "Columns", "SortOrder", .SortOrder
        
        For i = 1 To .ColumnHeaders.Count
            SaveSetting App.Title, "ColumnPos", format$(.ColumnHeaders(i).Position, "00"), i
            SaveSetting App.Title, "Columns", format$(i, "00"), .ColumnHeaders(i).Width
        Next
    End With
    
    SetOpstartDir Text1.Text

End
End Sub

Public Function FormatGenre(ByVal ID3Class As clsID3, ByVal GenreID As GenreConstants, ByVal Genre As String) As String
    If (GenreID = OtherGenre Or GenreID = Unknown) And Genre <> "" Then
        FormatGenre = Genre
    Else
        FormatGenre = ID3Class.GenreName(GenreID)
    End If
End Function

Public Function FormatTime(ByVal TimeVal As Double, Optional ByVal StoreTime As Boolean = False) As String
    On Error Resume Next
    
    Dim tv As Double
    Dim hr As Double
    Dim min As Double
    Dim sec As Double
    Dim ts As String
    
    tv = TimeVal
    If tv <= 0 Then
        If StoreTime Then dDuration = 0
    Else
        If StoreTime Then dDuration = tv
        
        tv = Fix(tv)
        min = Fix(tv / 60)
        sec = tv - 60 * min
        hr = Fix(min / 60)
        min = min - 60 * hr
        
        ts = ":" & format$(sec, "00")
        If hr > 0 Then
            ts = CStr(hr) & ":" & format$(min, "00") & ts
        Else
            ts = CStr(min) & ts
        End If
        
        FormatTime = ts
    End If
End Function

Public Function FormatBitRate(ByVal BitRate As Double, ByVal Encoding As EncodingEnum, Optional ByVal StoreBitRate As Boolean = False) As String
    On Error Resume Next
    
    Dim br As Double
    br = BitRate
    If br <= 0 Then
        If StoreBitRate Then dBitRate = 0
    Else
        If StoreBitRate Then dBitRate = br
        FormatBitRate = CStr(Fix(br / 1000)) '& " kbps " & IIf(Encoding = CBR, "CBR", "VBR")
    End If
End Function


Private Sub LoadFileEntries(ByVal Path As String)
    'On Error Resume Next
    Dim TrackNr As Integer
    Dim ID3 As New clsID3
    Dim sPath As String
    Dim d As String
    Dim HourPart As String
    Dim BlankWCOM As New MultiFrameData
    Dim BlankWOAR As New MultiFrameData
    Dim BlankAPIC As New MultiFrameData
    
    sPath = Path
    If Right$(Path, 1) <> "\" Then sPath = sPath & "\"
    
    d = Dir$(sPath)
    If ListView1.ListItems.Count < 1 Then
      ListView1.ListItems.Clear
      TrackNr = 0
    Else
    TrackNr = ListView1.ListItems.Count
    End If
    ID3Revision = 3
    'ShowOrHideNecessaryFields
    
    'ChangeFields False

    
    Do Until d = ""
        If d <> "." And d <> ".." Then
            If LCase$(Right$(d, 4)) = ".mp3" Then
                With ListView1
                    If MousePointer = vbDefault Then
                        MousePointer = vbHourglass
                        DoEvents
                    End If
                    FrmLoading.Label1.Caption = d
                    FrmLoading.Refresh
                    DoEvents
                    .ListItems.Add Text:=sPath & d
                    ID3.filename = sPath & d
                    TrackNr = TrackNr + 1
                    With .ListItems(.ListItems.Count)
                        .SubItems(1) = TrackNr & "."
                        .SubItems(2) = ID3.Title
                        .SubItems(3) = ID3.Artist
                        .SubItems(4) = ID3.Album
                        .SubItems(5) = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                        .SubItems(6) = ID3.TrackNumber
                        .SubItems(7) = ID3.TracksTotal
                        .SubItems(8) = ID3.Year
                        .SubItems(9) = FormatTime(ID3.Length)
                        .SubItems(10) = FormatBitRate(ID3.BitRate, ID3.Encoding)
                        .SubItems(11) = ID3.Comments
                    End With
                End With
            End If
        End If
        d = Dir$
    Loop
    
    Resort = True
    SortLvwOnLong ListView1, 2 'ListView1.SortKey + 1
    Resort = False
    
    If MousePointer = vbHourglass Then _
       MousePointer = vbDefault
    
    
    If ListView1.ListItems.Count > 0 Then
        'ChangeFields True
        ListView1.ListItems(1).Selected = True
        ListView1_ItemClick ListView1.ListItems(1)
    Else
        'ChangeFields False
    End If
    LblTotalTimeGenralList.Caption = CalcTotalTime(ListView1)
    Unload FrmLoading
End Sub


Private Sub AddFileEntries(ByVal filename As String, lv As ListView)
    'On Error Resume Next
    Dim TrackNr As Integer
    Dim ID3 As New clsID3

    

    
    ID3Revision = 3
    'ShowOrHideNecessaryFields
    
    'ChangeFields False

    TrackNr = lv.ListItems.Count
    With lv
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        FrmLoading.Label1.Caption = filename
        FrmLoading.Refresh
        DoEvents
        .ListItems.Add Text:=filename
        ID3.filename = filename
        TrackNr = TrackNr + 1
        With .ListItems(.ListItems.Count)
            .SubItems(1) = TrackNr & "."
            .SubItems(2) = ID3.Title
            .SubItems(3) = ID3.Artist
            .SubItems(4) = ID3.Album
            .SubItems(5) = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
            .SubItems(6) = ID3.TrackNumber
            .SubItems(7) = ID3.TracksTotal
            .SubItems(8) = ID3.Year
            .SubItems(9) = FormatTime(ID3.Length)
            .SubItems(10) = FormatBitRate(ID3.BitRate, ID3.Encoding)
            .SubItems(11) = ID3.Comments
        End With
    End With
    

    
    If MousePointer = vbHourglass Then _
       MousePointer = vbDefault
    
    
    If lv.ListItems.Count > 0 Then
        'ChangeFields True
        'lv.ListItems(1).Selected = True
        'ListView1_ItemClick ListView1.ListItems(1)
    Else
        'ChangeFields False
    End If
    Unload FrmLoading
End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  'PopupMenu mnuSystem
End If
End Sub

Private Sub ImgClose_Click()
  FrmshutDown.Show vbModal
  If FrmshutDown.ok = True Then
    Unload FrmshutDown
    Unload Me
  Else
     Unload FrmshutDown
  End If
 

End Sub

Private Sub Imgmin_Click()
Me.WindowState = 1
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim i As Long
    Dim idx As Long
    
    SortLvwOnLong ListView1, ColumnHeader.Index
    ShowListViewColumnHeaderSortIcon ListView1
    EnsureSelVisible ListView1
End Sub

Private Sub ListView1_DblClick()
    If SelectedIndex(ListView1) <> -1 Then
      LVmediaPLayer.URL = ListView1.selectedItem.Text
      LVmediaPLayer.Controls.play
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ID3 As New clsID3
    Dim HourPart As String
    Dim tempStr As String
    Dim sGenreID As String
    Dim bResort As Boolean
    Dim sItem As ListItem
    Dim bRefresh As Boolean
    Dim idx As Long
    
    bResort = False
    bRefresh = False
    Set sItem = ListView1.selectedItem
    
    If Dir$(sItem.Text) = "" Then
        ListView1.ListItems.Remove SelectedItemIdx(ListView1) + 1
        Exit Sub
    End If
    
    With ID3
        .filename = sItem.Text
        ID3Revision = .ID3RevisionV2
    End With
    
    With sItem
        If .SubItems(2) <> ID3.Title Then
            bRefresh = True
            If ListView1.SortKey = 2 Then bResort = True
            .SubItems(2) = ID3.Title
        End If
        
        If .SubItems(3) <> ID3.Artist Then
            bRefresh = True
            If ListView1.SortKey = 3 Then bResort = True
            .SubItems(3) = ID3.Artist
        End If
        
        If .SubItems(4) <> ID3.Album Then
            bRefresh = True
            If ListView1.SortKey = 4 Then bResort = True
            .SubItems(4) = ID3.Album
        End If
        
        tempStr = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
        If .SubItems(5) <> tempStr Then
            bRefresh = True
            If ListView1.SortKey = 5 Then bResort = True
            .SubItems(5) = tempStr
        End If
        
        If .SubItems(6) <> ID3.TrackNumber Then
            bRefresh = True
            If ListView1.SortKey = 6 Then bResort = True
            .SubItems(6) = ID3.TrackNumber
        End If
        
        If .SubItems(7) <> ID3.TracksTotal Then
            bRefresh = True
            If ListView1.SortKey = 7 Then bResort = True
            .SubItems(7) = ID3.TracksTotal
        End If
        
        If .SubItems(8) <> ID3.Year Then
            bRefresh = True
            If ListView1.SortKey = 8 Then bResort = True
            .SubItems(8) = ID3.Year
        End If
        
        tempStr = FormatTime(ID3.Length, True)
        If .SubItems(9) <> tempStr Then
            bRefresh = True
            If ListView1.SortKey = 9 Then bResort = True
            .SubItems(9) = tempStr
        End If
        
        tempStr = FormatBitRate(ID3.BitRate, ID3.Encoding, True)
        If .SubItems(10) <> tempStr Then
            bRefresh = True
            If ListView1.SortKey = 10 Then bResort = True
            .SubItems(10) = tempStr
        End If
        
        If .SubItems(11) <> ID3.Comments Then
            bRefresh = True
            If ListView1.SortKey = 11 Then bResort = True
            .SubItems(11) = ID3.Comments
        End If
    End With
    
    If bResort Then
        Resort = True
        SortLvwOnLong ListView1, ListView1.SortKey + 1
        Resort = False
    End If
    
    If bRefresh Then EnsureSelVisible ListView1, True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        ListView1_DblClick
    End If
End Sub


Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Exit Sub
Dim liNew       As ListItem
Dim pinfo       As LVHITTESTINFO
Dim pt          As POINTAPI
Dim pti         As POINTAPI
Dim hitItem     As ListItem
Dim i           As Integer
Dim bNew        As Boolean
   
On Error GoTo Err_Handler
Set hitItem = ListView1.HitTest(X, Y)

If Not hitItem Is Nothing Then

    Set liNew = ListView1.ListItems.Add(hitItem.Index, , DragLV.Text)
    i = 1
    Do Until i = DragLV.ListSubItems.Count + 1
        liNew.SubItems(i) = DragLV.SubItems(i)
        i = i + 1
    Loop
    liNew.Selected = True
    'listview1Copy.ListItems.Remove DragLV.Index
        
Else
    
    GetCursorPos pt
    
    If ListView1.ListItems.Count < 2 Then
        bNew = True
        Set liNew = ListView1.ListItems.Add(, , DragLV.Text)
        'Call ListView_GetItemPosition(listview1Copy.hWnd, _
            ListView1.ListItems.Item(ListView1.ListItems.Count).Index, pti)
    Else
    
        Call ListView_GetItemPosition(ListView1.hWnd, _
            ListView1.ListItems.Item(ListView1.ListItems.Count - 1).Index, pti)
    End If
    
    If pt.Y > Me.Top / Screen.TwipsPerPixelY + pti.Y Then
        If bNew = False Then Set liNew = ListView1.ListItems.Add(, , DragLV.Text)
            i = 1
            Do Until i = DragLV.ListSubItems.Count + 1
                liNew.SubItems(i) = DragLV.SubItems(i)
                i = i + 1
            Loop
            liNew.Selected = True
        ListView1.ListItems.Remove DragLV.Index
    End If
    
End If

Exit Sub
Err_Handler:

    
    
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & err & ": " & Error & " ", vbExclamation

End Sub



Private Sub ListView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Set DragLV = ListView1.selectedItem

End Sub

Private Sub LvPlaylist1_DblClick()
  Dim s As String
  Dim Index As Integer
  
  Index = LvPlaylist1.selectedItem.Index
  MorphLCDTrackA.Value = LvPlaylist1.ListItems(Index).SubItems(1)
  Deck1_File = LvPlaylist1.ListItems(Index).SubItems(2)

  Deck1.URL = LvPlaylist1.ListItems(Index).Text
  'Deck1_File.Caption = Mid(DragLV.Text, InStrRevVB5(DragLV.Text, "\") + 1, Len(DragLV.Text))
  Deck1_Remain.Caption = LvPlaylist1.ListItems(Index).SubItems(9)
  
  If Len(LvPlaylist1.ListItems(Index).SubItems(9)) = 4 Then Deck1_Remain.Caption = "00:0" & LvPlaylist1.ListItems(Index).SubItems(9)
  s = format(Deck1_Remain.Caption, "hh:mm:ss")
  If Left(s, 1) = "0" Then
     s = Right(s, Len(s) - 1)
     MorphLCDRemainingTimeA.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
     MorphLCDRemainingTimeA.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
  MorphLCDElapsedTimeA.Value = "000000"
    

End Sub
Private Sub orderPLaylist1()
  Dim i As Integer
    For i = 1 To LvPlaylist1.ListItems.Count
      LvPlaylist1.ListItems(i).SubItems(1) = i
    Next i
End Sub

Private Sub orderPLaylist2()
  Dim i As Integer
    For i = 1 To LvPlaylist2.ListItems.Count
      LvPlaylist2.ListItems(i).SubItems(1) = i
    Next i
End Sub
Private Sub LvPlaylist1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ID3 As New clsID3
    Dim HourPart As String
    Dim tempStr As String
    Dim sGenreID As String
    Dim bResort As Boolean
    Dim sItem As ListItem
    Dim bRefresh As Boolean
    Dim idx As Long
    
    bResort = False
    bRefresh = False
    Set sItem = LvPlaylist1.selectedItem
    
    If Dir$(sItem.Text) = "" Then
        LvPlaylist1.ListItems.Remove SelectedItemIdx(LvPlaylist1) + 1
        Exit Sub
    End If
    
    With ID3
        .filename = sItem.Text
        ID3Revision = .ID3RevisionV2
    End With
    
    With sItem
        If .SubItems(2) <> ID3.Title Then
            bRefresh = True
            If LvPlaylist1.SortKey = 2 Then bResort = True
            .SubItems(2) = ID3.Title
        End If
        
        If .SubItems(3) <> ID3.Artist Then
            bRefresh = True
            If LvPlaylist1.SortKey = 3 Then bResort = True
            .SubItems(3) = ID3.Artist
        End If
        
        If .SubItems(4) <> ID3.Album Then
            bRefresh = True
            If LvPlaylist1.SortKey = 4 Then bResort = True
            .SubItems(4) = ID3.Album
        End If
        
        tempStr = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
        If .SubItems(5) <> tempStr Then
            bRefresh = True
            If LvPlaylist1.SortKey = 5 Then bResort = True
            .SubItems(5) = tempStr
        End If
        
        If .SubItems(6) <> ID3.TrackNumber Then
            bRefresh = True
            If LvPlaylist1.SortKey = 6 Then bResort = True
            .SubItems(6) = ID3.TrackNumber
        End If
        
        If .SubItems(7) <> ID3.TracksTotal Then
            bRefresh = True
            If LvPlaylist1.SortKey = 7 Then bResort = True
            .SubItems(7) = ID3.TracksTotal
        End If
        
        If .SubItems(8) <> ID3.Year Then
            bRefresh = True
            If LvPlaylist1.SortKey = 8 Then bResort = True
            .SubItems(8) = ID3.Year
        End If
        
        tempStr = FormatTime(ID3.Length, True)
        If .SubItems(9) <> tempStr Then
            bRefresh = True
            If LvPlaylist1.SortKey = 9 Then bResort = True
            .SubItems(9) = tempStr
        End If
        
        tempStr = FormatBitRate(ID3.BitRate, ID3.Encoding, True)
        If .SubItems(10) <> tempStr Then
            bRefresh = True
            If LvPlaylist1.SortKey = 10 Then bResort = True
            .SubItems(10) = tempStr
        End If
        
        If .SubItems(11) <> ID3.Comments Then
            bRefresh = True
            If LvPlaylist1.SortKey = 11 Then bResort = True
            .SubItems(11) = ID3.Comments
        End If
    End With
    
    If bResort Then
        Resort = True
        SortLvwOnLong LvPlaylist1, LvPlaylist1.SortKey + 1
        Resort = False
    End If
    
    If bRefresh Then EnsureSelVisible LvPlaylist1, True
End Sub

Private Sub LvPlaylist1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim liNew       As ListItem
'Dim pinfo       As LVHITTESTINFO
Dim pt          As POINTAPI
Dim pti         As POINTAPI
Dim hitItem     As ListItem
Dim i           As Integer
Dim bNew        As Boolean
Dim s As String


On Error GoTo Err_Handler

Set hitItem = LvPlaylist1.HitTest(X, Y)

    Set liNew = LvPlaylist1.ListItems.Add(, , DragLV.Text)
    i = 1
    Do Until i = DragLV.ListSubItems.Count + 1
        liNew.SubItems(i) = DragLV.SubItems(i)
        i = i + 1
    Loop
    liNew.SubItems(1) = LvPlaylist1.ListItems.Count
    'liNew.Selected = True
    'ListView1.ListItems.Remove DragLV.Index
    If LvPlaylist1.ListItems.Count < 2 Then
      MorphLCDTrackA.Value = LvPlaylist1.ListItems.Count
      Deck1_File = liNew.SubItems(2)
      Deck1.URL = DragLV.Text
      'Deck1_File.Caption = Mid(DragLV.Text, InStrRevVB5(DragLV.Text, "\") + 1, Len(DragLV.Text))
      Deck1_Remain.Caption = liNew.SubItems(9)
      
      If Len(liNew.SubItems(9)) = 4 Then Deck1_Remain.Caption = "00:0" & liNew.SubItems(9)
      s = format(Deck1_Remain.Caption, "hh:mm:ss")
      If Left(s, 1) = "0" Then
         s = Right(s, Len(s) - 1)
         MorphLCDRemainingTimeA.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
      Else
         MorphLCDRemainingTimeA.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
      End If
      MorphLCDElapsedTimeA.Value = "000000"
    End If
    LblTotalTimePlaylist1.Caption = CalcTotalTime(LvPlaylist1)
    Call orderPLaylist1
Exit Sub
Err_Handler:
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & err & ": " & Error & " ", vbExclamation

End Sub



Private Sub LvPlaylist2_DblClick()
  Dim s As String
  Dim Index As Integer
  
  Index = LvPlaylist2.selectedItem.Index
  MorphLCDTrackB.Value = LvPlaylist2.ListItems(Index).SubItems(1)
  Deck2_File = LvPlaylist2.ListItems(Index).SubItems(2)
  Deck2.URL = LvPlaylist2.ListItems(Index).Text
  'Deck1_File.Caption = Mid(DragLV.Text, InStrRevVB5(DragLV.Text, "\") + 1, Len(DragLV.Text))
  Deck2_Remain.Caption = LvPlaylist2.ListItems(Index).SubItems(9)
  
  If Len(LvPlaylist2.ListItems(Index).SubItems(9)) = 4 Then Deck2_Remain.Caption = "00:0" & LvPlaylist2.ListItems(Index).SubItems(9)
  s = format(Deck2_Remain.Caption, "hh:mm:ss")
  If Left(s, 1) = "0" Then
     s = Right(s, Len(s) - 1)
     MorphLCDRemainingTimeB.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
     MorphLCDRemainingTimeB.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
  MorphLCDElapsedTimeB.Value = "000000"
End Sub

Private Sub LvPlaylist2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim ID3 As New clsID3
    Dim HourPart As String
    Dim tempStr As String
    Dim sGenreID As String
    Dim bResort As Boolean
    Dim sItem As ListItem
    Dim bRefresh As Boolean
    Dim idx As Long
    
    bResort = False
    bRefresh = False
    Set sItem = LvPlaylist2.selectedItem
    
    If Dir$(sItem.Text) = "" Then
        LvPlaylist2.ListItems.Remove SelectedItemIdx(LvPlaylist2) + 1
        Exit Sub
    End If
    
    With ID3
        .filename = sItem.Text
        ID3Revision = .ID3RevisionV2
    End With
    
    With sItem
        If .SubItems(2) <> ID3.Title Then
            bRefresh = True
            If LvPlaylist2.SortKey = 2 Then bResort = True
            .SubItems(2) = ID3.Title
        End If
        
        If .SubItems(3) <> ID3.Artist Then
            bRefresh = True
            If LvPlaylist2.SortKey = 3 Then bResort = True
            .SubItems(3) = ID3.Artist
        End If
        
        If .SubItems(4) <> ID3.Album Then
            bRefresh = True
            If LvPlaylist2.SortKey = 4 Then bResort = True
            .SubItems(4) = ID3.Album
        End If
        
        tempStr = FormatGenre(ID3, ID3.GenreID, ID3.Genre)
        If .SubItems(5) <> tempStr Then
            bRefresh = True
            If LvPlaylist2.SortKey = 5 Then bResort = True
            .SubItems(5) = tempStr
        End If
        
        If .SubItems(6) <> ID3.TrackNumber Then
            bRefresh = True
            If LvPlaylist2.SortKey = 6 Then bResort = True
            .SubItems(6) = ID3.TrackNumber
        End If
        
        If .SubItems(7) <> ID3.TracksTotal Then
            bRefresh = True
            If LvPlaylist2.SortKey = 7 Then bResort = True
            .SubItems(7) = ID3.TracksTotal
        End If
        
        If .SubItems(8) <> ID3.Year Then
            bRefresh = True
            If LvPlaylist2.SortKey = 8 Then bResort = True
            .SubItems(8) = ID3.Year
        End If
        
        tempStr = FormatTime(ID3.Length, True)
        If .SubItems(9) <> tempStr Then
            bRefresh = True
            If LvPlaylist2.SortKey = 9 Then bResort = True
            .SubItems(9) = tempStr
        End If
        
        tempStr = FormatBitRate(ID3.BitRate, ID3.Encoding, True)
        If .SubItems(10) <> tempStr Then
            bRefresh = True
            If LvPlaylist2.SortKey = 10 Then bResort = True
            .SubItems(10) = tempStr
        End If
        
        If .SubItems(11) <> ID3.Comments Then
            bRefresh = True
            If LvPlaylist2.SortKey = 11 Then bResort = True
            .SubItems(11) = ID3.Comments
        End If
    End With
    
    If bResort Then
        Resort = True
        SortLvwOnLong LvPlaylist2, LvPlaylist2.SortKey + 1
        Resort = False
    End If
    
    If bRefresh Then EnsureSelVisible LvPlaylist2, True

End Sub

Private Sub LvPlaylist2_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim liNew       As ListItem
Dim pinfo       As LVHITTESTINFO
Dim pt          As POINTAPI
Dim pti         As POINTAPI
Dim hitItem     As ListItem
Dim i           As Integer
Dim bNew        As Boolean
   Dim s As String
On Error GoTo Err_Handler

Set hitItem = LvPlaylist2.HitTest(X, Y)


    Set liNew = LvPlaylist2.ListItems.Add(, , DragLV.Text)
    i = 1
    Do Until i = DragLV.ListSubItems.Count + 1
        liNew.SubItems(i) = DragLV.SubItems(i)
        i = i + 1
    Loop
    liNew.SubItems(1) = LvPlaylist2.ListItems.Count
    'liNew.Selected = True
    'ListView1.ListItems.Remove DragLV.Index
    If LvPlaylist2.ListItems.Count < 2 Then
      MorphLCDTrackB.Value = LvPlaylist2.ListItems.Count
      Deck2_File = liNew.SubItems(2)
      Deck2.URL = DragLV.Text
      'Deck2_File.Caption = Mid(DragLV.Text, InStrRevVB5(DragLV.Text, "\") + 1, Len(DragLV.Text))
      Deck2_Remain.Caption = liNew.SubItems(9)
      
      If Len(liNew.SubItems(9)) = 4 Then Deck2_Remain.Caption = "00:0" & liNew.SubItems(9)
      s = format(Deck2_Remain.Caption, "hh:mm:ss")
      If Left(s, 1) = "0" Then
         s = Right(s, Len(s) - 1)
         MorphLCDRemainingTimeB.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
      Else
         MorphLCDRemainingTimeB.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
      End If
      MorphLCDElapsedTimeB.Value = "000000"
    End If
    LblTotalTimePlaylist2.Caption = CalcTotalTime(LvPlaylist2)
    Call orderPLaylist2
Exit Sub
Err_Handler:
        
    MsgBox "An error has occurred! " & vbCrLf & vbCrLf & err & ": " & Error & " ", vbExclamation
End Sub

Private Sub LvPLayPlay_Click()
  If SelectedIndex(ListView1) <> -1 Then
      LVmediaPLayer.URL = ListView1.selectedItem.Text
      LVmediaPLayer.Controls.play
  End If
End Sub

Private Sub LvStopPlay_Click()
  LVmediaPLayer.Controls.stop
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub MSSlider1_ValueHasChanged()
Cross_Fader.Value = MSSlider1.Value
End Sub

Private Sub MSSlider2_ValueHasChanged()
Deck1_Volume.Value = MSSlider2.Value
End Sub

Private Sub MSSlider3_ValueHasChanged()
Deck2_Volume.Value = MSSlider3.Value
End Sub

Private Sub MSSlider4_ValueHasChanged()
SampPlayer.settings.volume = MSSlider4.Value
End Sub

Private Sub Option1_Click()
'Slider1.Value = -100
'Slider2.Value = 100
'Slider1.Enabled = False
'Slider2.Enabled = False
End Sub

Private Sub Option2_Click()
'Slider1.Value = 0
'Slider2.Value = 0
'Slider1.Enabled = True
'Slider2.Enabled = True
End Sub



Private Sub Show_File_Finder_Click()
On Error Resume Next
'Form2.Show
'Form2.Top = FrmSearch.Top + FrmSearch.Height
'Form2.Left = FrmSearch.Left
End Sub



Private Sub PicAutoFade_Click()
If ChkAutomaticMixing.Value = vbChecked Then
  ChkAutomaticMixing.Value = vbUnchecked
  PicautoFade.Picture = Frmpics.PicautoFade(0).Image
Else
  ChkAutomaticMixing.Value = vbChecked
  PicautoFade.Picture = Frmpics.PicautoFade(1).Image
End If
End Sub

Private Sub Picbtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicBtnTmp.Picture = Picbtn(Index).Image
  Picbtn(Index).Picture = PicbtnDown(Index).Image
End Sub

Private Sub Picbtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim itmx As ListItem
Dim i As Long
  Picbtn(Index).Picture = PicBtnTmp.Image
  Select Case Index
    Case Is = 0
      ListView1.ListItems.Clear
    Case Is = 1
        Dim Folder As String
        Dim sExistingFolder As String
        
        If Right$(Text1, 1) = "\" Then
            sExistingFolder = Text1
        Else
            sExistingFolder = Text1 & "\"
        End If
        
        Folder = BrowseForFolder(hWnd, "Select a folder:", sExistingFolder)
        If Folder <> "" Then
            Text1 = Folder
            LoadFileEntries Folder
        End If
    Case Is = 2
      FrmFileDialog.Show vbModal
      For i = 0 To List1.ListCount - 1
        AddFileEntries List1.list(i), ListView1
      Next i
      List1.Clear
    Case Is = 7
      Resort = True
      SortLvwOnLong ListView1, 3 'ListView1.SortKey + 1
      Resort = False
    Case Is = 8
        If ListView1.ListItems.Count < 1 Then Exit Sub
        FrmMp3TagInfo.filename = ListView1.selectedItem
        FrmMp3TagInfo.Show vbModal
        ListView1_ItemClick ListView1.selectedItem
    Case Is = 9
      FrmSearch.Show vbModal
      ListView1.ListItems.Clear
      For i = 0 To List1.ListCount - 1
        AddFileEntries List1.list(i), ListView1
      Next i
      List1.Clear
      LblTotalTimeGenralList.Caption = CalcTotalTime(ListView1)
    Case Is = 10
        If LvPlaylist1.ListItems.Count < 1 Then Exit Sub
        FrmMp3TagInfo.filename = LvPlaylist1.selectedItem
        FrmMp3TagInfo.Show vbModal
        LvPlaylist1_ItemClick LvPlaylist1.selectedItem
    Case Is = 12
      If LvPlaylist1.ListItems.Count < 1 Then Exit Sub
      LvPlaylist1.ListItems.Remove (LvPlaylist1.selectedItem.Index)
      LblTotalTimePlaylist1.Caption = CalcTotalTime(LvPlaylist1)
      Call orderPLaylist1
    Case Is = 13
       For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
          AddFileEntries ListView1.ListItems(i).Text, LvPlaylist1
        End If
      Next i
      LblTotalTimePlaylist1.Caption = CalcTotalTime(LvPlaylist1)
      Call orderPLaylist1
    Case Is = 14 'playlist1 save
      With Dialog
        '// Reset file name
        .filename = ""
        '// Change dialog title
        .DialogTitle = "Save M3U"
        '// Set filter
        .Filter = "PLS Playlists (*.m3u)|*.m3u"
        '// Show open dialog
        .ShowSave
        '// Check for blank file name
        If IsBlank(.filename) Then Exit Sub
        '// Save M3U
        SaveM3U .filename, LvPlaylist1
      End With
    Case Is = 15 'playlist1 load
      With Dialog
        '// Reset file name
        .filename = ""
        '// Change dialog title
        .DialogTitle = "Load M3U"
        '// Set filter
        .Filter = "M3U Playlists (*.m3u)|*.m3u"
        '// Show open dialog
        .ShowOpen
        '// Check for blank file name
        If IsBlank(.filename) Then Exit Sub
        '// Load M3U
        LoadM3U .filename, LvPlaylist1
        LblTotalTimePlaylist1.Caption = CalcTotalTime(LvPlaylist1)
        
      End With
    Case Is = 16
      LvPlaylist1.ListItems.Clear
    Case Is = 17
      LvPlaylist2.ListItems.Clear
    Case Is = 18 'playlist1 load
      With Dialog
        '// Reset file name
        .filename = ""
        '// Change dialog title
        .DialogTitle = "Load M3U"
        '// Set filter
        .Filter = "M3U Playlists (*.m3u)|*.m3u"
        '// Show open dialog
        .ShowOpen
        '// Check for blank file name
        If IsBlank(.filename) Then Exit Sub
        '// Load M3U
        LoadM3U .filename, LvPlaylist2
        LblTotalTimePlaylist2.Caption = CalcTotalTime(LvPlaylist2)
      End With
    Case Is = 19 'playlist1 save
      With Dialog
        '// Reset file name
        .filename = ""
        '// Change dialog title
        .DialogTitle = "Save M3U"
        '// Set filter
        .Filter = "PLS Playlists (*.m3u)|*.m3u"
        '// Show open dialog
        .ShowSave
        '// Check for blank file name
        If IsBlank(.filename) Then Exit Sub
        '// Save M3U
        SaveM3U .filename, LvPlaylist2
      End With
    Case Is = 20
       For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
          AddFileEntries ListView1.ListItems(i).Text, LvPlaylist2
        End If
      Next i
      LblTotalTimePlaylist2.Caption = CalcTotalTime(LvPlaylist2)
      Call orderPLaylist2
    Case Is = 21
      If LvPlaylist2.ListItems.Count < 1 Then Exit Sub
      LvPlaylist2.ListItems.Remove (LvPlaylist2.selectedItem.Index)
      LblTotalTimePlaylist2.Caption = CalcTotalTime(LvPlaylist2)
      Call orderPLaylist2
    Case Is = 23
      If LvPlaylist2.ListItems.Count < 1 Then Exit Sub
      FrmMp3TagInfo.filename = LvPlaylist2.selectedItem
      FrmMp3TagInfo.Show vbModal
      LvPlaylist2_ItemClick LvPlaylist2.selectedItem
    Case Is = 24
      On Local Error Resume Next
      err.Clear
      
      Shell "sndvol32.exe", vbNormalFocus
    Case Is = 25
      FrmAbout.Show vbModal
  End Select
End Sub

Private Sub PicbtnSamplePlayer_Click(Index As Integer)

  Select Case Index
    Case Is = 0
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path1", App.Path & "\wav\ELECTRIC A.wav")
    Case Is = 1
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path2", App.Path & "\wav\PECHE A.wav")
    Case Is = 2
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path3", App.Path & "\wav\WAVE_AHAHBPM.wav")
    Case Is = 3
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path4", App.Path & "\wav\FX BONGA.wav")
    Case Is = 4
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path5", App.Path & "\wav\WOOSHE A.wav")
    Case Is = 5
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path6", App.Path & "\wav\WOOSHE B.wav")
    Case Is = 6
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path7", App.Path & "\wav\LOOP 140 ELECTRO 05.wav")
    Case Is = 7
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path8", App.Path & "\wav\BPM.wav")
    Case Is = 8
      SampPlayer.URL = GetSetting(App.Title, "Samplersound", "Path9", App.Path & "\wav\LOOP 185 ANALOG 01.wav")
  End Select
  SampPlayer.Controls.play
End Sub

Private Sub PicDnPlaylist1_Click()
ListViewMoveSelDown LvPlaylist1
End Sub

Private Sub PicEditSampler_Click()
FrmSamplerEditor.Show vbModal
End Sub

Private Sub PicFadeA_Click()
  On Error Resume Next
  'If timerfade.Value = 1 Then
  Deck1.Controls.play
  timerfade2.Enabled = False
  timerfade1.Enabled = True
  'Else
  'Cross_Fader.Value = Cross_Fader.Value - 5
  'timerfade2.Enabled = False
  'timerfade1.Enabled = False
  'End If
End Sub

Private Sub PicFadeB_Click()
  On Error Resume Next
  'If timerfade1.Enabled = 1 Then
  Deck2.Controls.play
  timerfade2.Enabled = True
  timerfade1.Enabled = False
  'Else
  'Cross_Fader.Value = Cross_Fader.Value + 5
  'timerfade2.Enabled = False
  'timerfade1.Enabled = False
  'End If
End Sub

Private Sub PicFind_Click()
    Resort = True
    SortLvwOnLong ListView1, 2 'ListView1.SortKey + 1
    Resort = False

  FindLVItem ListView1, TxtFind.Text ', , tmpMultiSelect, tmpInverseSelection

End Sub

Private Sub PicFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicFind.Picture = Frmpics.PicFind(1).Image
End Sub

Private Sub PicFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicFind.Picture = Frmpics.PicFind(0).Image
End Sub

Private Sub PicFindNext_Click()
  FindLVItem ListView1, TxtFind.Text, , , , True ', , tmpMultiSelect, tmpInverseSelection

End Sub

Private Sub PicFindNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicFindNext.Picture = Frmpics.PicFindNext(1).Image
End Sub

Private Sub PicFindNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicFindNext.Picture = Frmpics.PicFindNext(0).Image
End Sub

Private Sub PicPLayer1Loop_Click()
If ChkLoopPLayer1.Value = vbChecked Then
  ChkLoopPLayer1.Value = vbUnchecked
  PicPLayer1Loop.Picture = Frmpics.PicLoop(0).Image
Else
  ChkLoopPLayer1.Value = vbChecked
  PicPLayer1Loop.Picture = Frmpics.PicLoop(1).Image
End If
End Sub

Private Sub PicPlayer1Shuffle_Click()

If ChkRandomPLayer1.Value = vbChecked Then
  ChkRandomPLayer1.Value = vbUnchecked
  PicPlayer1Shuffle.Picture = Frmpics.PicPlayerShuffle(0).Image
Else
  ChkRandomPLayer1.Value = vbChecked
  PicPlayer1Shuffle.Picture = Frmpics.PicPlayerShuffle(1).Image
End If
End Sub

Private Sub PicPLayer1SinglePLay_Click()
If ChkSinglePLayer1.Value = vbChecked Then
  ChkSinglePLayer1.Value = vbUnchecked
  PicPLayer1SinglePLay.Picture = Frmpics.PicSingleplay(0).Image
Else
  ChkSinglePLayer1.Value = vbChecked
  PicPLayer1SinglePLay.Picture = Frmpics.PicSingleplay(1).Image
End If
End Sub

Private Sub PicPLayer2Loop_Click()
If ChkLoopPLayer2.Value = vbChecked Then
  ChkLoopPLayer2.Value = vbUnchecked
  PicPLayer2Loop.Picture = Frmpics.PicLoop(0).Image
Else
  ChkLoopPLayer2.Value = vbChecked
  PicPLayer2Loop.Picture = Frmpics.PicLoop(1).Image
End If
End Sub

Private Sub PicPlayer2Shuffle_Click()
If ChkRandomPLayer2.Value = vbChecked Then
  ChkRandomPLayer2.Value = vbUnchecked
  PicPlayer2Shuffle.Picture = Frmpics.PicPlayerShuffle(0).Image
Else
  ChkRandomPLayer2.Value = vbChecked
  PicPlayer2Shuffle.Picture = Frmpics.PicPlayerShuffle(1).Image
End If

End Sub

Private Sub PicPLayer2SinglePLay_Click()
If ChkSinglePLayer2.Value = vbChecked Then
  ChkSinglePLayer2.Value = vbUnchecked
  PicPLayer2SinglePLay.Picture = Frmpics.PicSingleplay(0).Image
Else
  ChkSinglePLayer2.Value = vbChecked
  PicPLayer2SinglePLay.Picture = Frmpics.PicSingleplay(1).Image
End If
End Sub

Private Sub PicPlayPL1_Click()
  Dim s As String
  Dim Index As Integer
  On Error Resume Next
  Index = LvPlaylist1.selectedItem.Index
  If err Then Exit Sub
  
  If Deck1Pause = False Then
    Deck1.Controls.pause
    'Timerdeck1Pause.Enabled = True
    Deck1Pause = True
  Else
    'Timerdeck1Pause.Enabled = False
    PicLightPlayPL1.Visible = True
    PicLightStopPL1.Visible = False
    Deck1Pause = False
  
    GoTo go
    
    MorphLCDTrackA.Value = LvPlaylist1.ListItems(Index).SubItems(1)
    Deck1_File = LvPlaylist1.ListItems(Index).SubItems(2)
    Deck1.URL = LvPlaylist1.ListItems(Index).Text
    'Deck1_File.Caption = Mid(DragLV.Text, InStrRevVB5(DragLV.Text, "\") + 1, Len(DragLV.Text))
    Deck1_Remain.Caption = LvPlaylist1.ListItems(Index).SubItems(9)
    
    If Len(LvPlaylist1.ListItems(Index).SubItems(9)) = 4 Then Deck1_Remain.Caption = "00:0" & LvPlaylist1.ListItems(Index).SubItems(9)
    s = format(Deck1_Remain.Caption, "hh:mm:ss")
    If Left(s, 1) = "0" Then
       s = Right(s, Len(s) - 1)
       MorphLCDRemainingTimeA.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
    Else
       MorphLCDRemainingTimeA.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
    End If
    MorphLCDElapsedTimeA.Value = "000000"
go:
    Deck1.Controls.play
  End If
End Sub


Private Sub PicPlayPL2_Click()
  Dim s As String
  Dim Index As Integer
  
  On Error Resume Next
  Index = LvPlaylist2.selectedItem.Index
  If err Then Exit Sub
  
  If Deck2Pause = False Then
    Deck2.Controls.pause
    'Timerdeck1Pause.Enabled = True
    Deck2Pause = True
  Else
  
  'Timerdeck1Pause.Enabled = False
  PicLightPlayPL2.Visible = True
  PicLightStopPL2.Visible = False
  Deck2Pause = False
  GoTo go

  MorphLCDTrackB.Value = LvPlaylist2.ListItems(Index).SubItems(1)
  Deck2_File = LvPlaylist2.ListItems(Index).SubItems(2)
  Deck2.URL = LvPlaylist2.ListItems(Index).Text
  'Deck1_File.Caption = Mid(DragLV.Text, InStrRevVB5(DragLV.Text, "\") + 1, Len(DragLV.Text))
  Deck2_Remain.Caption = LvPlaylist2.ListItems(Index).SubItems(9)
  
  If Len(LvPlaylist2.ListItems(Index).SubItems(9)) = 4 Then Deck2_Remain.Caption = "00:0" & LvPlaylist2.ListItems(Index).SubItems(9)
  s = format(Deck2_Remain.Caption, "hh:mm:ss")
  If Left(s, 1) = "0" Then
     s = Right(s, Len(s) - 1)
     MorphLCDRemainingTimeB.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
     MorphLCDRemainingTimeB.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
  MorphLCDElapsedTimeB.Value = "000000"
go:
  Deck2.Controls.play
End If
End Sub

Private Sub PicResetFader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicResetFader.Picture = Frmpics.PicResetfading(1).Image
End Sub

Private Sub PicResetFader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicResetFader.Picture = Frmpics.PicResetfading(0).Image
End Sub

Private Sub PicStopPL1_Click()
  Deck1.Controls.stop
  ECScrollingText1.Scrolling = False

  PicLightPlayPL1.Visible = False
  PicLightStopPL1.Visible = True = True
  PlayerAStopped = True
  Deck1Pause = True
  'Timerdeck1Pause.Enabled = False
End Sub

Private Sub PicPlayPL1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicPlayPL1.Picture = Frmpics.PicbtnPlay(1).Image
End Sub

Private Sub PicPlayPL1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicPlayPL1.Picture = Frmpics.PicbtnPlay(0).Image
End Sub

Private Sub PicPlayPL2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicPlayPL2.Picture = Frmpics.PicbtnPlay(1).Image
End Sub

Private Sub PicPlayPL2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicPlayPL2.Picture = Frmpics.PicbtnPlay(0).Image
End Sub

Private Sub PicStopPL1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicStopPL1.Picture = Frmpics.PicbtnStop(1).Image
End Sub

Private Sub PicStopPL1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicStopPL1.Picture = Frmpics.PicbtnStop(0).Image
End Sub

Private Sub PicStopPL2_Click()
  Deck2.Controls.stop
  
  ECScrollingText2.Scrolling = False

  PicLightPlayPL2.Visible = False
  PicLightStopPL2.Visible = True = True
  Deck2Pause = True
  PlayerBStopped = True
  'Timerdeck1Pause.Enabled = False

End Sub

Private Sub PicStopPL2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicStopPL2.Picture = Frmpics.PicbtnStop(1).Image
End Sub

Private Sub PicStopPL2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PicStopPL2.Picture = Frmpics.PicbtnStop(0).Image
End Sub

Private Sub PicTitlebar_Click()
Unload Me
End Sub

Private Sub PicTrackLeftPlayerA_Click()
  Dim xx As Integer
  Dim xlast As Integer
  Dim Index  As Integer
  On Error Resume Next
  Index = LvPlaylist1.selectedItem.Index
  If err Then Exit Sub
  
  Index = LvPlaylist1.selectedItem.Index
  If LvPlaylist1.selectedItem.Index = 1 Then
      LvPlaylist1.ListItems(1).Selected = True
      xx = 1
      GoTo go
  End If
  xx = LvPlaylist1.selectedItem.Index - 1
go:
          
  LvPlaylist1.ListItems(xx).Selected = True
  MorphLCDTrackA.Value = LvPlaylist1.ListItems(xx).SubItems(1)
  Deck1_File = LvPlaylist1.ListItems(xx).SubItems(2)
  Deck1.URL = LvPlaylist1.ListItems(xx).Text
  Deck1_Remain.Caption = LvPlaylist1.ListItems(xx).SubItems(9)
  
  Deck1.Controls.play
End Sub

Private Sub PicTrackLeftPlayerB_Click()
  Dim xx As Integer
  Dim xlast As Integer
  Dim Index  As Integer
  
  On Error Resume Next
  Index = LvPlaylist1.selectedItem.Index
  If err Then Exit Sub
  Index = LvPlaylist2.selectedItem.Index
  If LvPlaylist2.selectedItem.Index = 1 Then
      LvPlaylist2.ListItems(1).Selected = True
      xx = 1
      GoTo go
  End If
  xx = LvPlaylist2.selectedItem.Index - 1
go:
          
  LvPlaylist2.ListItems(xx).Selected = True
  MorphLCDTrackB.Value = LvPlaylist2.ListItems(xx).SubItems(1)
  Deck2_File = LvPlaylist2.ListItems(xx).SubItems(2)
  Deck2.URL = LvPlaylist2.ListItems(xx).Text
  Deck2_Remain.Caption = LvPlaylist2.ListItems(xx).SubItems(9)
  
  Deck2.Controls.play
End Sub

Private Sub PicTrackRightPlayerA_Click()
  Dim xx As Integer
  Dim xlast As Integer
  Dim Index  As Integer
  
  On Error Resume Next
  Index = LvPlaylist1.selectedItem.Index
  If err Then Exit Sub
  PlayerAStopped = True
  xx = LvPlaylist1.selectedItem.Index + 1
  If LvPlaylist1.selectedItem.Index = LvPlaylist1.ListItems.Count Then
      LvPlaylist1.ListItems(1).Selected = True
      xx = 1
  End If
          
  LvPlaylist1.ListItems(xx).Selected = True
  MorphLCDTrackA.Value = LvPlaylist1.ListItems(xx).SubItems(1)
  Deck1_File = LvPlaylist1.ListItems(xx).SubItems(2)
  Deck1.URL = LvPlaylist1.ListItems(xx).Text
  Deck1_Remain.Caption = LvPlaylist1.ListItems(xx).SubItems(9)
  
  Deck1.Controls.play
  PlayerAStopped = False
End Sub

Private Sub PicTrackRightPlayerB_Click()
  Dim xx As Integer
  Dim xlast As Integer
  Dim Index  As Integer
  
  On Error Resume Next
  Index = LvPlaylist1.selectedItem.Index
  If err Then Exit Sub
  PlayerBStopped = True
  Index = LvPlaylist2.selectedItem.Index
  xx = LvPlaylist2.selectedItem.Index + 1
  If LvPlaylist2.selectedItem.Index = LvPlaylist2.ListItems.Count Then
      LvPlaylist2.ListItems(1).Selected = True
      xx = 1
  End If
          
  LvPlaylist2.ListItems(xx).Selected = True
  MorphLCDTrackB.Value = LvPlaylist2.ListItems(xx).SubItems(1)
  Deck2_File = LvPlaylist2.ListItems(xx).SubItems(2)
  Deck2.URL = LvPlaylist2.ListItems(xx).Text
  Deck2_Remain.Caption = LvPlaylist2.ListItems(xx).SubItems(9)
  
  Deck2.Controls.play
  PlayerBStopped = False
End Sub

Private Sub PicResetFader_Click()
  timerfade1.Enabled = False
  timerfade2.Enabled = False
  Cross_Fader.Value = 0
  PicResetFader.Enabled = False
End Sub

Private Sub PicUpPlaylist1_Click()
ListViewMoveSelUp LvPlaylist1
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  Dim s As String
  Dim xx, xlast As Long
  
  
  If (Deck1.playState = wmppsStopped) Then
    Index = LvPlaylist1.selectedItem.Index
    xlast = LvPlaylist1.selectedItem.Index
    
    If LvPlaylist1.selectedItem.Index = LvPlaylist1.ListItems.Count Then
      LvPlaylist1.ListItems(1).Selected = True
      xx = LvPlaylist1.selectedItem.Index
      GoTo go
    End If
    If ChkRandomPLayer1.Value = vbChecked Then
      Do Until xx <> xlast And xx > 0
      DoEvents
          Randomize
          xx = Rnd * LvPlaylist1.ListItems.Count
      Loop
    Else
      xx = LvPlaylist1.selectedItem.Index + 1
    End If
go:
    LvPlaylist1.ListItems(xx).Selected = True
    MorphLCDTrackA.Value = LvPlaylist1.ListItems(xx).SubItems(1)
    Deck1_File = LvPlaylist1.ListItems(xx).SubItems(2)
    Deck1.URL = LvPlaylist1.ListItems(xx).Text
    Deck1_Remain.Caption = LvPlaylist1.ListItems(xx).SubItems(9)
    

    If ChkAutomaticMixing.Value = vbUnchecked Then
      Deck1.Controls.play
      PlayerAStopped = False
    End If
  
  End If
  
  If (Deck2.playState = wmppsStopped) Then
    Index = LvPlaylist2.selectedItem.Index
    xlast = LvPlaylist2.selectedItem.Index
    If LvPlaylist2.selectedItem.Index = LvPlaylist2.ListItems.Count Then
        LvPlaylist2.ListItems(1).Selected = True
        xx = LvPlaylist2.selectedItem.Index
        GoTo go2
    End If
    If ChkRandomPLayer2.Value = vbChecked Then
        Do Until xx <> xlast And xx > 0
        DoEvents
            Randomize
            xx = Rnd * LvPlaylist2.ListItems.Count
        Loop
    Else
      xx = LvPlaylist2.selectedItem.Index + 1
    End If
go2:
    LvPlaylist2.ListItems(xx).Selected = True
    MorphLCDTrackB.Value = LvPlaylist2.ListItems(xx).SubItems(1)
    Deck1_File = LvPlaylist2.ListItems(xx).SubItems(2)
    Deck1.URL = LvPlaylist2.ListItems(xx).Text
    Deck2_Remain.Caption = LvPlaylist2.ListItems(xx).SubItems(9)
    

    If ChkAutomaticMixing.Value = vbUnchecked Then
      Deck2.Controls.play
      PlayerBStopped = False
    End If
End If

' Show time
' > DECK 1
On Error Resume Next
If Deck1.Controls.currentPosition > 0 Then
  Deck1_Time.Caption = TimeSerial(0, 0, Int(Deck1.Controls.currentPosition))
  Label23.Caption = Deck1_Time.Caption
  s = format(TimeSerial(0, 0, Int(Deck1.Controls.currentPosition)), "hh:mm:ss")
  If Left(s, 1) = "0" Then
    s = Right(s, Len(s) - 1)
    MorphLCDElapsedTimeA.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
    MorphLCDElapsedTimeA.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
  Randomize
  For i = 0 To 8 Step 1
    sp(i).Height = Rnd() * (300)
    sp(i).Top = (255 + 285) - sp(i).Height
  Next

'Remaining time
  Deck1_Remain.Caption = "" & TimeSerial(0, 0, Int(Deck1.currentMedia.duration) - Int(Deck1.Controls.currentPosition)) & ""

  s = "000000"
  s = format("" & TimeSerial(0, 0, Int(Deck1.currentMedia.duration) - Int(Deck1.Controls.currentPosition)) & "", "hh:mm:ss")
  If Left(s, 1) = "0" Then
    s = Right(s, Len(s) - 1)
    MorphLCDRemainingTimeA.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
    MorphLCDRemainingTimeA.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
End If
' > DECK 2
On Error Resume Next
If Deck2.Controls.currentPosition > 0 Then
  Deck2_Time.Caption = TimeSerial(0, 0, Int(Deck2.Controls.currentPosition))

  s = format(TimeSerial(0, 0, Int(Deck2.Controls.currentPosition)), "hh:mm:ss")
  If Left(s, 1) = "0" Then
    s = Right(s, Len(s) - 1)
    MorphLCDElapsedTimeB.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
    MorphLCDElapsedTimeB.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
  Randomize
  For i = 9 To 17 Step 1
    sp(i).Height = Rnd() * (300)
    sp(i).Top = (255 + 285) - sp(i).Height
  Next

'Remaining time
  Deck2_Remain.Caption = "" & TimeSerial(0, 0, Int(Deck2.currentMedia.duration) - Int(Deck2.Controls.currentPosition)) & ""

  s = "000000"
  s = format("" & TimeSerial(0, 0, Int(Deck2.currentMedia.duration) - Int(Deck2.Controls.currentPosition)) & "", "hh:mm:ss")
  If Left(s, 1) = "0" Then
    s = Right(s, Len(s) - 1)
    MorphLCDRemainingTimeB.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
  Else
    MorphLCDRemainingTimeB.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
  End If
End If

' Turn mp3 name to red if 20 seconds or less left in track
' DECK 1
If Deck1.Controls.currentPosition >= (Deck1.currentMedia.duration - 20) Then
  MorphLCDRemainingTimeA.SegmentLitColor = vbRed
Else
  MorphLCDRemainingTimeA.SegmentLitColor = &HFFFF00
End If

' DECK 2
If Deck2.Controls.currentPosition >= (Deck2.currentMedia.duration - 20) Then
  MorphLCDRemainingTimeB.SegmentLitColor = vbRed
Else
  MorphLCDRemainingTimeB.SegmentLitColor = &HFFFF00
End If

ECSlider1.Value = Deck1.Controls.currentPosition
ECSlider1.Max = Deck1.currentMedia.duration

ECSlider2.Value = Deck2.Controls.currentPosition
ECSlider2.Max = Deck2.currentMedia.duration


'Debug.Print SoundMeter.getVolume(buffaddress)
'LevelOfSound.Value = SoundMeter.getVolume(buffaddress)
'drawScope

'LevelOfSound1.Cls
'LevelOfSound1.Line (0, (LevelOfSound1.ScaleHeight))-(LevelOfSound1.Width, (LevelOfSound1.ScaleHeight - SoundMeter.getVolume(buffaddress))), vbCyan, BF
'If SoundMeter.getVolume(buffaddress) > 1 Then
'P_Val = SoundMeter.getVolume(buffaddress)
'drawmeter
'End If
End Sub


Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar1.Value = Deck1_Volume.Value + 100
ProgressBar2.Value = Deck2_Volume.Value + 100

If ChkAutomaticMixing.Value = 1 And Deck1_Remain.Caption = "0:00:15" Then
  timerfade2.Enabled = True
  timerfade1.Enabled = False
  Deck2.Controls.play
Else
End If

If ChkAutomaticMixing.Value = 1 And Deck2_Remain.Caption = "0:00:15" Then
  timerfade2.Enabled = False
  timerfade1.Enabled = True
  Deck1.Controls.play
Else
End If
End Sub

Private Sub Timer3_Timer()
If Cross_Fader.Value <> 0 Then
  PicResetFader.Enabled = True
Else
  PicResetFader.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
'Show_File_Finder_Click
'Timer4.Enabled = False
End Sub


Private Sub Timerdeck1Pause_Timer()
  PicLightPlayPL1.Visible = Not PicLightPlayPL1.Visible
End Sub

Private Sub Timerdeck2Pause_Timer()
  PicLightPlayPL2.Visible = Not PicLightPlayPL2.Visible

End Sub

Private Sub timerfade1_Timer()
On Error Resume Next
Cross_Fader.Value = Cross_Fader.Value - 5
End Sub

Private Sub timerfade2_Timer()
On Error Resume Next
Cross_Fader.Value = Cross_Fader.Value + 5
End Sub
Sub ProcessStream(Stream As String)
OldBPM = Asc(Left$(Stream, 1)) + Asc(Right$(Stream, 1)) * 256
End Sub
Function GetLatest() As String
GetLatest = Chr$(Val(Label2.Caption)) + OutBuffer + Chr$(0)
OutBuffer = ""
End Function
Function InitializeDevice(id As Byte) As String
Dim TempID As Long
TempID = Device_Channel + Channel_ChannelID + Channel_Commands + Channel_BPM
'Debug.Print TempID
InitializeDevice = Chr$(TempID And &HFF&) + Chr$((TempID And &HFF00&) / &H100&) + Chr$((TempID And &HFF0000) / &H10000) + Chr$(0)

End Function
Private Sub Command7_Click()
On Error Resume Next
Static LastClick
If LastClick <> 0 And LastClick < Timer Then
    LastBPM = (LastBPM + 1) Mod 15
    If MaxBPMs < 15 Then MaxBPMs = MaxBPMs + 1
    BPMArray(LastBPM) = 60 / (Timer - LastClick)
    For i = 0 To 14
        cBPM = cBPM + BPMArray(i)
    Next
    Label8.Caption = format(cBPM / MaxBPMs, "##0.00")
    BPM = (cBPM / MaxBPMs) * 100
End If

LastClick = Timer
End Sub



Private Sub VSlider1_Change(Value As Long)
Deck1_Volume.Value = Value
End Sub

Private Sub VSlider2_Change(Value As Long)
Deck2_Volume.Value = Value
End Sub



Private Sub ListViewMoveSelUp(ByVal lv As ListView)
    Dim bWasUnSel As Boolean
    Dim tmpLvItem As ListItem
    Dim newLvItem As ListItem
    Dim tmpSubItem As ListSubItem
    Dim i As Integer
    bWasUnSel = False
    For i = 1 To lv.ListItems.Count
        Set tmpLvItem = lv.ListItems(i)
        If tmpLvItem.Selected Then
            If bWasUnSel Then
                Set newLvItem = lv.ListItems.Add(tmpLvItem.Index - 1, , tmpLvItem.Text)
                For Each tmpSubItem In tmpLvItem.ListSubItems
                    newLvItem.SubItems(tmpSubItem.Index) = tmpSubItem.Text
                Next
                lv.ListItems.Remove (tmpLvItem.Index)
                newLvItem.Selected = True
                Set newLvItem = Nothing
            End If
        Else
            bWasUnSel = True
        End If
        Set tmpLvItem = Nothing
    Next
    For i = 1 To lv.ListItems.Count
      lv.ListItems(i).SubItems(1) = i
    Next i
End Sub
Private Sub ListViewMoveSelDown(ByVal lv As ListView)
    Dim bWasUnSel As Boolean
    Dim tmpLvItem As ListItem
    Dim newLvItem As ListItem
    Dim tmpSubItem As ListSubItem
    Dim i As Integer
    bWasUnSel = False
    For i = lv.ListItems.Count To 1 Step -1
        Set tmpLvItem = lv.ListItems(i)
        If tmpLvItem.Selected Then
            If bWasUnSel Then
                Set newLvItem = lv.ListItems.Add(tmpLvItem.Index + 2, , tmpLvItem.Text)
                For Each tmpSubItem In tmpLvItem.ListSubItems
                    newLvItem.SubItems(tmpSubItem.Index) = tmpSubItem.Text
                Next
                lv.ListItems.Remove (tmpLvItem.Index)
                newLvItem.Selected = True
                Set newLvItem = Nothing
            End If
        Else
            bWasUnSel = True
        End If
        Set tmpLvItem = Nothing
    Next
        For i = 1 To lv.ListItems.Count
      lv.ListItems(i).SubItems(1) = i
    Next i
    
End Sub
