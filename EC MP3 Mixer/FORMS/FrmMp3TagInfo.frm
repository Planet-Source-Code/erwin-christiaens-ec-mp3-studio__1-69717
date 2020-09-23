VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMp3TagInfo 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Split"
      Height          =   315
      Left            =   9045
      TabIndex        =   98
      Top             =   60
      Width           =   600
   End
   Begin VB.TextBox TxtFilename 
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
      Left            =   120
      TabIndex        =   96
      Top             =   90
      Width           =   8760
   End
   Begin VB.Frame Frame1 
      Caption         =   "MP3 ID Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   9735
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   390
         Left            =   5175
         TabIndex        =   97
         Top             =   3600
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   9255
         Begin VB.TextBox txtTitle 
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
            Left            =   600
            TabIndex        =   10
            Top             =   0
            Width           =   8655
         End
         Begin VB.TextBox txtArtist 
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
            Left            =   600
            TabIndex        =   9
            Top             =   360
            Width           =   8655
         End
         Begin VB.TextBox txtAlbum 
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
            Left            =   600
            TabIndex        =   8
            Top             =   720
            Width           =   8655
         End
         Begin VB.ComboBox cmbGenre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmMp3TagInfo.frx":0000
            Left            =   600
            List            =   "FrmMp3TagInfo.frx":01C3
            TabIndex        =   7
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtTrackNumber 
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
            Left            =   4800
            TabIndex        =   6
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtComments 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtYear 
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
            Left            =   6960
            TabIndex        =   4
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtLyrics 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox txtTracksTotal 
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
            Left            =   5640
            TabIndex        =   2
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Title:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Artist:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Album:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Genre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Track:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4260
            TabIndex        =   15
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Comments:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Year:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   13
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Lyrics:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   12
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label39 
            Caption         =   "of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   11
            Top             =   1080
            Width           =   255
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Update MP3 Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   94
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete MP3 Tags"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   93
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   2
         Left            =   240
         TabIndex        =   85
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.ComboBox cmbPictureType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmMp3TagInfo.frx":0798
            Left            =   3960
            List            =   "FrmMp3TagInfo.frx":07DB
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   2280
            Width           =   2520
         End
         Begin VB.PictureBox picArt 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   3960
            ScaleHeight     =   1755
            ScaleWidth      =   1755
            TabIndex        =   87
            Top             =   0
            Width           =   1815
            Begin VB.Label lblBrowse 
               Alignment       =   2  'Center
               BackColor       =   &H8000000C&
               BackStyle       =   0  'Transparent
               Caption         =   "Click here to browse..."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   495
               Left            =   0
               TabIndex        =   88
               Top             =   720
               Width           =   1815
            End
            Begin VB.Image imgArt 
               Height          =   1755
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1755
            End
         End
         Begin VB.ComboBox cmbImageType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmMp3TagInfo.frx":090A
            Left            =   3960
            List            =   "FrmMp3TagInfo.frx":091A
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   1920
            Width           =   1275
         End
         Begin VB.Label Label41 
            Caption         =   "Image type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   92
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label43 
            Caption         =   "Picture type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   91
            Top             =   2280
            Width           =   975
         End
         Begin VB.Image prevArt 
            Height          =   285
            Left            =   5850
            Picture         =   "FrmMp3TagInfo.frx":0933
            Top             =   1920
            Width           =   210
         End
         Begin VB.Image nextArt 
            Height          =   285
            Left            =   6060
            Picture         =   "FrmMp3TagInfo.frx":0A13
            Top             =   1920
            Width           =   210
         End
         Begin VB.Image delArt 
            Height          =   285
            Left            =   6270
            Picture         =   "FrmMp3TagInfo.frx":0AF2
            Top             =   1920
            Width           =   210
         End
         Begin VB.Label countArt 
            Alignment       =   1  'Right Justify
            Caption         =   "0/0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   5265
            TabIndex        =   90
            Top             =   1980
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   9255
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   9825
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   9015
            Begin VB.TextBox txtEncodedBy 
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
               Left            =   1800
               TabIndex        =   52
               Top             =   9000
               Width           =   7095
            End
            Begin VB.TextBox txtPaymentURL 
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
               Left            =   1800
               TabIndex        =   51
               Top             =   8280
               Width           =   7095
            End
            Begin VB.TextBox txtAudioURL 
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
               Left            =   1800
               TabIndex        =   50
               Top             =   6840
               Width           =   7095
            End
            Begin VB.ComboBox cmbKey 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "FrmMp3TagInfo.frx":0BD3
               Left            =   2880
               List            =   "FrmMp3TagInfo.frx":0C43
               TabIndex        =   49
               Top             =   9480
               Width           =   1455
            End
            Begin VB.TextBox txtBPM 
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
               Left            =   480
               TabIndex        =   48
               Top             =   9480
               Width           =   1095
            End
            Begin VB.TextBox txtLyricist 
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
               Left            =   1800
               TabIndex        =   47
               Top             =   1440
               Width           =   7095
            End
            Begin VB.TextBox txtLanguages 
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
               Left            =   1800
               TabIndex        =   46
               Top             =   5760
               Width           =   7095
            End
            Begin VB.TextBox txtCopyright 
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
               Left            =   1800
               TabIndex        =   45
               Top             =   3600
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalReleaseYear 
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
               Left            =   1800
               TabIndex        =   44
               Top             =   3240
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalLyricist 
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
               Left            =   1800
               TabIndex        =   43
               Top             =   2880
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalFileName 
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
               Left            =   1800
               TabIndex        =   42
               Top             =   2520
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalAlbum 
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
               Left            =   1800
               TabIndex        =   41
               Top             =   2160
               Width           =   7095
            End
            Begin VB.TextBox txtOriginalArtist 
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
               Left            =   1800
               TabIndex        =   40
               Top             =   1800
               Width           =   7095
            End
            Begin VB.TextBox txtComposer 
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
               Left            =   1800
               TabIndex        =   39
               Top             =   0
               Width           =   7095
            End
            Begin VB.TextBox txtPublisherURL 
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
               Left            =   1800
               TabIndex        =   38
               Top             =   8640
               Width           =   7095
            End
            Begin VB.TextBox txtBand 
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
               Left            =   1800
               TabIndex        =   37
               Top             =   360
               Width           =   7095
            End
            Begin VB.TextBox txtConductor 
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
               Left            =   1800
               TabIndex        =   36
               Top             =   720
               Width           =   7095
            End
            Begin VB.TextBox txtFileOwner 
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
               Left            =   1800
               TabIndex        =   35
               Top             =   3960
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationName 
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
               Left            =   1800
               TabIndex        =   34
               Top             =   4680
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationOwner 
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
               Left            =   1800
               TabIndex        =   33
               Top             =   5040
               Width           =   7095
            End
            Begin VB.TextBox txtISRC 
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
               Left            =   1800
               TabIndex        =   32
               Top             =   5400
               Width           =   7095
            End
            Begin VB.TextBox txtCommercialInfo 
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
               Left            =   1800
               TabIndex        =   31
               Top             =   6120
               Width           =   5850
            End
            Begin VB.TextBox txtCopyrightInfo 
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
               Left            =   1800
               TabIndex        =   30
               Top             =   6480
               Width           =   7095
            End
            Begin VB.TextBox txtArtistURL 
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
               Left            =   1800
               TabIndex        =   29
               Top             =   7200
               Width           =   5850
            End
            Begin VB.TextBox txtAudioSourceURL 
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
               Left            =   1800
               TabIndex        =   28
               Top             =   7560
               Width           =   7095
            End
            Begin VB.TextBox txtInternetRadioStationURL 
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
               Left            =   1800
               TabIndex        =   27
               Top             =   7920
               Width           =   7095
            End
            Begin VB.TextBox txtDiscNumber 
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
               Left            =   5160
               TabIndex        =   26
               Top             =   9480
               Width           =   495
            End
            Begin VB.TextBox txtDiscsTotal 
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
               Left            =   6000
               TabIndex        =   25
               Top             =   9480
               Width           =   495
            End
            Begin VB.TextBox txtPublisher 
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
               Left            =   1800
               TabIndex        =   24
               Top             =   4320
               Width           =   7095
            End
            Begin VB.TextBox txtInterpretedBy 
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
               Left            =   1800
               TabIndex        =   23
               Top             =   1080
               Width           =   7095
            End
            Begin VB.Label Label23 
               Caption         =   "Encoded by:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   84
               Top             =   9000
               Width           =   1695
            End
            Begin VB.Label Label22 
               Caption         =   "Payment URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   83
               Top             =   8280
               Width           =   1695
            End
            Begin VB.Label Label21 
               Caption         =   "Audio URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   82
               Top             =   6840
               Width           =   1695
            End
            Begin VB.Label Label20 
               Caption         =   "Initial Key:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1920
               TabIndex        =   81
               Top             =   9480
               Width           =   855
            End
            Begin VB.Label Label19 
               Caption         =   "BPM:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   80
               Top             =   9480
               Width           =   375
            End
            Begin VB.Label Label17 
               Caption         =   "Lyricist:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   79
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label16 
               Caption         =   "Languages:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   78
               Top             =   5760
               Width           =   1695
            End
            Begin VB.Label Label15 
               Caption         =   "Copyright:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   77
               Top             =   3600
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Original Release Year:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   76
               Top             =   3240
               Width           =   1695
            End
            Begin VB.Label Label13 
               Caption         =   "Original Lyricist:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   75
               Top             =   2880
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Original Filename:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   74
               Top             =   2520
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "Original Album:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   73
               Top             =   2160
               Width           =   1695
            End
            Begin VB.Label Label10 
               Caption         =   "Original Artist:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   72
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label Label9 
               Caption         =   "Composer:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   71
               Top             =   0
               Width           =   1695
            End
            Begin VB.Label Label24 
               Caption         =   "Publisher URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   70
               Top             =   8640
               Width           =   1695
            End
            Begin VB.Label Label25 
               Caption         =   "Band/Orchestra:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   69
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label26 
               Caption         =   "Conductor:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   68
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label27 
               Caption         =   "File Owner/Licensee:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   67
               Top             =   3960
               Width           =   1695
            End
            Begin VB.Label Label28 
               Caption         =   "Net Radio Stn. Name:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   66
               Top             =   4680
               Width           =   1695
            End
            Begin VB.Label Label29 
               Caption         =   "Net Radio Stn. Owner:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   65
               Top             =   5040
               Width           =   1695
            End
            Begin VB.Label Label30 
               Caption         =   "ISRC:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   64
               Top             =   5400
               Width           =   1695
            End
            Begin VB.Label Label31 
               Caption         =   "Commercial Info URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   63
               Top             =   6120
               Width           =   1695
            End
            Begin VB.Label Label32 
               Caption         =   "Copyright Info URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   62
               Top             =   6480
               Width           =   1695
            End
            Begin VB.Label Label33 
               Caption         =   "Artist URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   61
               Top             =   7200
               Width           =   1695
            End
            Begin VB.Label Label34 
               Caption         =   "Audio Source URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   60
               Top             =   7560
               Width           =   1695
            End
            Begin VB.Label Label35 
               Caption         =   "Net Radio Station URL:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   59
               Top             =   7920
               Width           =   1695
            End
            Begin VB.Label Label36 
               Caption         =   "Disc:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4680
               TabIndex        =   58
               Top             =   9480
               Width           =   375
            End
            Begin VB.Label Label37 
               Caption         =   "of"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5760
               TabIndex        =   57
               Top             =   9480
               Width           =   255
            End
            Begin VB.Label Label38 
               Caption         =   "Publisher:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   56
               Top             =   4320
               Width           =   1695
            End
            Begin VB.Label Label40 
               Caption         =   "Interpreted by:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   55
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Image prevCommercialInfo 
               Height          =   285
               Left            =   8265
               Picture         =   "FrmMp3TagInfo.frx":0DFB
               ToolTipText     =   "Previous Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image nextCommercialInfo 
               Height          =   285
               Left            =   8475
               Picture         =   "FrmMp3TagInfo.frx":0EDB
               ToolTipText     =   "Next Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image delCommercialInfo 
               Height          =   285
               Left            =   8685
               Picture         =   "FrmMp3TagInfo.frx":0FBA
               ToolTipText     =   "Delete Commercial Info URL"
               Top             =   6120
               Width           =   210
            End
            Begin VB.Image prevArtistURL 
               Height          =   285
               Left            =   8265
               Picture         =   "FrmMp3TagInfo.frx":109B
               ToolTipText     =   "Previous Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image nextArtistURL 
               Height          =   285
               Left            =   8475
               Picture         =   "FrmMp3TagInfo.frx":117B
               ToolTipText     =   "Next Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Image delArtistURL 
               Height          =   285
               Left            =   8685
               Picture         =   "FrmMp3TagInfo.frx":125A
               ToolTipText     =   "Delete Artist URL"
               Top             =   7200
               Width           =   210
            End
            Begin VB.Label countCommercialInfo 
               Alignment       =   1  'Right Justify
               Caption         =   "0/0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   7680
               TabIndex        =   54
               Top             =   6180
               Width           =   570
            End
            Begin VB.Label countArtistURL 
               Alignment       =   1  'Right Justify
               Caption         =   "0/0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   7680
               TabIndex        =   53
               Top             =   7260
               Width           =   570
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   2655
            LargeChange     =   5
            Left            =   9000
            Max             =   29
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3255
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5741
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Basic"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Advanced"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Album Art"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList Buttons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":133B
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":13B2
            Key             =   "addi"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":1430
            Key             =   "del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":1521
            Key             =   "deli"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":1612
            Key             =   "next"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":1701
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMp3TagInfo.frx":17F1
            Key             =   "previ"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMp3TagInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Const SW_SHOW As Long = 5
Private Const CF_BITMAP As Long = 2
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4

Private Const S_OTHER As String = "Other"

Private Const FILTER_BMP As String = "*.bmp;*.dib"
Private Const FILTER_GIF As String = "*.gif"
Private Const FILTER_JPEG As String = "*.jpeg;*.jpg;*.jpe;*.jfif;*.jfi;*.jif"
Private Const FILTER_PNG As String = "*.png"
Private Const FILTER_SUPPORTED As String = FILTER_BMP & ";" & FILTER_GIF & ";" & FILTER_JPEG & ";" & FILTER_PNG


Public filename As String

Private Sub AddAPICItem(ByVal MIMEType As String, ByVal PictureType As PictureType, ByVal Data As String)
    Dim APD As APicDecoder
    
    If Data = "" Then
        cAPICIType.Add ""
        cAPICType.Add ""
        cAPICData.Add ""
    Else
        cAPICIType.Add MIMEType
        cAPICType.Add PictureType
        cAPICData.Add Data
    End If
    APICData.Add ""
    
    If Data <> "" Then
        Set APD = New APicDecoder
        APD.InsertImageData APICData, APICData.Count, MIMEType, PictureType, Data, ID3Revision
        Set APD = Nothing
    End If
    
    cAPIC0.Add APICData.Count
End Sub

Private Function FormatGenre(ByVal ID3Class As clsID3, ByVal GenreID As GenreConstants, ByVal Genre As String) As String
    If (GenreID = OtherGenre Or GenreID = Unknown) And Genre <> "" Then
        FormatGenre = Genre
    Else
        FormatGenre = ID3Class.GenreName(GenreID)
    End If
End Function

Private Function FormatTime(ByVal TimeVal As Double, Optional ByVal StoreTime As Boolean = False) As String
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

Private Function FormatBitRate(ByVal BitRate As Double, ByVal Encoding As EncodingEnum, Optional ByVal StoreBitRate As Boolean = False) As String
    On Error Resume Next
    
    Dim br As Double
    br = BitRate
    If br <= 0 Then
        If StoreBitRate Then dBitRate = 0
    Else
        If StoreBitRate Then dBitRate = br
        FormatBitRate = CStr(Fix(br / 1000)) & " kbps " & IIf(Encoding = CBR, "CBR", "VBR")
    End If
End Function

Private Sub cmbImageType_Change()
    cmbImageType_Click
End Sub

Private Sub cmbImageType_Click()
    Dim MIMEType As String
    Dim PNGIndex As Long
    If cmbImageType.Enabled Then
        If cmbImageType.ListCount = 4 Then
            Select Case cmbImageType.ListIndex
                Case 0: MIMEType = ImageBMP
                Case 1: MIMEType = ImageGIF
                Case 2: MIMEType = ImageJPEG
                Case 3: MIMEType = ImagePNG
            End Select
            PNGIndex = 3
        Else
            Select Case cmbImageType.ListIndex
                Case 0: MIMEType = ImageJPEGOld
                Case 1: MIMEType = ImagePNGOld
            End Select
            PNGIndex = 1
        End If
        SetItem cAPICIType, indAPIC, MIMEType
        If cmbPictureType.ListIndex = 1 And cmbImageType.ListIndex <> PNGIndex Then
            cmbPictureType.ListIndex = 2
            SetItem cAPICType, indAPIC, cmbImageType.ListIndex
        End If
    End If
End Sub

Private Sub cmbPictureType_Change()
    cmbPictureType_Click
End Sub

Private Sub cmbPictureType_Click()
    If cmbPictureType.Enabled Then
        If cmbPictureType.ListIndex = 1 Then
            If cmbImageType.ListIndex <> (1 + 2 * (cmbImageType.ListCount \ 4)) Or HimetricToPixelsX(imgArt.Picture.Width) <> 32 Or HimetricToPixelsY(imgArt.Picture.Height) <> 32 Then
                cmbPictureType.ListIndex = 2
            End If
        End If
        SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
    End If
End Sub


Private Sub CmdClose_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
    Dim Path As Variant
    Dim strArray() As String
    Dim StripFileName As String
    Dim FullPath As String
    If InStr(TxtFilename, ".mp3") Then
      FullPath = Left(TxtFilename, Len(TxtFilename) - 4)
    Else
      FullPath = TxtFilename
    End If
    Path = Split(FullPath, "\")
    StripFileName = Path(UBound(Path))
    On Error Resume Next
    strArray = Split(StripFileName, "-")
    txtArtist = RTrim(LTrim(strArray(0)))
    txtTitle = RTrim(LTrim(strArray(1)))
End Sub


Private Sub Command2_Click()

    Dim ID3 As New clsID3
    Dim i As Long
    
    With ID3
        .filename = filename
        .Title = txtTitle
        .Artist = txtArtist
        .Album = txtAlbum
        .Genre = cmbGenre.Text
        .GenreID = .ToGenreID(.Genre)
        .TrackNumber = txtTrackNumber
        .TracksTotal = txtTracksTotal
        .Year = txtYear
        .Comments = txtComments
        .Lyrics = txtLyrics
        .Composer = txtComposer
        .Band = txtBand
        .Conductor = txtConductor
        .InterpretedBy = txtInterpretedBy
        .Lyricist = txtLyricist
        .OriginalArtist = txtOriginalArtist
        .OriginalAlbum = txtOriginalAlbum
        .OriginalFileName = txtOriginalFileName
        .OriginalLyricist = txtOriginalLyricist
        .OriginalReleaseYear = txtOriginalReleaseYear
        .Copyright = txtCopyright
        .FileOwner = txtFileOwner
        .Publisher = txtPublisher
        .InternetRadioStationName = txtInternetRadioStationName
        .InternetRadioStationOwner = txtInternetRadioStationOwner
        .ISRC = txtISRC
        .Languages = txtLanguages
        .CommercialInfo.Clear
        For i = 1 To cWCOM.Count
            .CommercialInfo.Add cWCOM(i)
        Next
        .CopyrightInfo = txtCopyrightInfo
        .AudioURL = txtAudioURL
        .ArtistURL.Clear
        For i = 1 To cWOAR.Count
            .ArtistURL.Add cWOAR(i)
        Next
        .AudioSourceURL = txtAudioSourceURL
        .InternetRadioURL = txtInternetRadioStationURL
        .PaymentURL = txtPaymentURL
        .PublisherURL = txtPublisherURL
        .EncodedBy = txtEncodedBy
        .BeatsPerMinute = txtBPM
        .InitialKey = cmbKey
        .DiscNumber = txtDiscNumber
        .DiscsTotal = txtDiscsTotal
        For i = 1 To cAPICData.Count
            MakeNecessaryChanges i
        Next
        Set .AttachedPictures = APICData
        
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        
        .UpdateID3Tags
        
        If MousePointer = vbHourglass Then _
           MousePointer = vbDefault
        'Debug.Print Form5.ListView1.selectedItem
        'Form5.ListView1_ItemClick Form5.ListView1.selectedItem
    End With
End Sub

Private Sub MakeNecessaryChanges(ByVal Index As Long)
    Dim APD As APicDecoder
    Dim GPC As GDIPlusCandy
    
    Dim MIMEType As String
    Dim PictureType As PictureType
    Dim Pic As StdPicture
    Dim PicData As String
    
    Set APD = New APicDecoder
    APD.DecodeImage APICData, cAPIC0(Index), MIMEType, PictureType, Pic, ID3Revision
    If MIMEType = cAPICIType(Index) Then
        PicData = cAPICData(Index)
    Else
        Set GPC = New GDIPlusCandy
        PicData = GPC.ImageToData(Pic, cAPICIType(Index))
        Set GPC = Nothing
    End If
    APD.InsertImageData APICData, cAPIC0(Index), cAPICIType(Index), cAPICType(Index), PicData, ID3Revision
    Set APD = Nothing
End Sub


Private Sub Command3_Click()
    Dim ID3 As New clsID3
    
    With ID3
        .filename = filename
        
        If MousePointer = vbDefault Then
            MousePointer = vbHourglass
            DoEvents
        End If
        
        .DeleteID3Tags
        
        If MousePointer = vbHourglass Then _
           MousePointer = vbDefault
        
        'Form5.ListView1_ItemClick Form5.ListView1.selectedItem
    End With
End Sub

Private Sub Form_Load()
 Frame2(0).Visible = True
 Frame2(1).Visible = False
 Frame2(2).Visible = False
 
    Dim ID3 As New clsID3
    Dim HourPart As String
    Dim tempStr As String
    Dim sGenreID As String

    Dim idx As Long
    
  TxtFilename = filename
    
    With ID3
        .filename = filename
        ID3Revision = .ID3RevisionV2
        'ShowOrHideNecessaryFields
        txtTitle = .Title
        txtArtist = .Artist
        txtAlbum = .Album
        cmbGenre = FormatGenre(ID3, .GenreID, .Genre)
        txtTrackNumber = .TrackNumber
        txtTracksTotal = .TracksTotal
        txtYear = .Year
        txtComments = .Comments
        txtLyrics = .Lyrics
        txtComposer = .Composer
        txtBand = .Band
        txtConductor = .Conductor
        txtInterpretedBy = .InterpretedBy
        txtLyricist = .Lyricist
        txtOriginalArtist = .OriginalArtist
        txtOriginalAlbum = .OriginalAlbum
        txtOriginalFileName = .OriginalFileName
        txtOriginalLyricist = .OriginalLyricist
        txtOriginalReleaseYear = .OriginalReleaseYear
        txtCopyright = .Copyright
        txtFileOwner = .FileOwner
        txtPublisher = .Publisher
        txtInternetRadioStationName = .InternetRadioStationName
        txtInternetRadioStationOwner = .InternetRadioStationOwner
        txtISRC = .ISRC
        txtLanguages = .Languages
        LoadMultiData txtArtistURL, .ArtistURL, S_AURL, countArtistURL, prevArtistURL, nextArtistURL, delArtistURL, cWOAR, indWOAR, totWOAR, bWOARBlank
        txtCopyrightInfo = .CopyrightInfo
        txtAudioURL = .AudioURL
        LoadMultiData txtCommercialInfo, .CommercialInfo, S_CURL, countCommercialInfo, prevCommercialInfo, nextCommercialInfo, delCommercialInfo, cWCOM, indWCOM, totWCOM, bWCOMBlank
        txtAudioSourceURL = .AudioSourceURL
        txtInternetRadioStationURL = .InternetRadioURL
        txtPaymentURL = .PaymentURL
        txtPublisherURL = .PublisherURL
        txtEncodedBy = .EncodedBy
        txtBPM = .BeatsPerMinute
        cmbKey = .InitialKey
        txtDiscNumber = .DiscNumber
        txtDiscsTotal = .DiscsTotal
        LoadMultiData picArt, .AttachedPictures, S_APIC, countArt, prevArt, nextArt, delArt, cAPICData, indAPIC, totAPIC, bAPICBlank
    End With
 
 
End Sub

Private Sub lblBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If WithinBounds(picArt, X + lblBrowse.Left, Y + lblBrowse.Top) Then
        Select Case Button
            Case vbLeftButton: ImageBrowse
            Case vbRightButton: If ValidateMenu Then PopupMenu mnuArt
        End Select
    End If
End Sub

Private Sub TabStrip1_Click()
    On Error Resume Next
    Dim i As Long
    For i = 1 To TabStrip1.Tabs.Count
        If Frame2(i - 1).Visible <> TabStrip1.Tabs(i).Selected Then
            Frame2(i - 1).Visible = TabStrip1.Tabs(i).Selected
        End If
    Next
End Sub

Private Sub VScroll1_Change()
    Dim FTop As Single: FTop = -CSng(VScroll1.Value) * 360
    Dim FTopMax As Single: FTopMax = -Frame3.Height + Frame2(1).Height
    
    If (VScroll1.Value = VScroll1.Max And FTop > FTopMax) Or FTop < FTopMax Then
        If Frame3.Top <> FTopMax Then Frame3.Top = FTopMax
    Else
        If Frame3.Top <> FTop Then Frame3.Top = FTop
    End If
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub


Private Function FilterEntry(ByVal Description As String, ByVal Filter As String) As String
    FilterEntry = Description & "|" & Filter & "|"
End Function


Private Sub ImageBrowse()
    Dim fn As String
    Dim f As Integer
    Dim st As String
    Dim sMIMEType As String
    Dim tMIMEType As String
    Dim sExt As String
    Dim GPC As GDIPlusCandy
    Dim sPic As StdPicture
    Dim i As Long
    Dim idx As Long
    Dim bConvertImage As Boolean
    
    'If ListView1.ListItems.Count > 0 Then
        fn = ShowOpenDialog(hWnd, FilterEntry("All Supported Formats", FILTER_SUPPORTED) & FilterEntry("Windows Bitmap", FILTER_BMP) & FilterEntry("Graphics Interchange Format", FILTER_GIF) & FilterEntry("JPEG File Interchange Format", FILTER_JPEG) & FilterEntry("Portable Network Graphics", FILTER_PNG), "Select Image")
        If fn <> "" Then
            i = InStrRev(fn, ".")
            If i > 0 Then
                sExt = Mid$(LCase$(fn), i + 1)
                Select Case sExt
                    Case "bmp", "dib"
                        sMIMEType = ImageTypeFromIndex(0, ID3Revision)
                        idx = 1
                        If ID3Revision > 2 Then idx = 0
                        bConvertImage = (ID3Revision <= 2)
                    Case "gif"
                        sMIMEType = ImageTypeFromIndex(1, ID3Revision)
                        idx = 1
                        bConvertImage = (ID3Revision <= 2)
                    Case "jpeg", "jpg", "jpe", "jfif", "jfi", "jif"
                        sMIMEType = ImageTypeFromIndex(2, ID3Revision)
                        idx = 0
                        If ID3Revision > 2 Then idx = 2
                    Case "png"
                        sMIMEType = ImageTypeFromIndex(3, ID3Revision)
                        idx = 1
                        If ID3Revision > 2 Then idx = 3
                    Case Else
                        sMIMEType = ""
                        idx = -1
                End Select
                If idx <> -1 Then cmbImageType.ListIndex = idx
            Else
                sMIMEType = ""
            End If
            
            f = FreeFile
            Open fn For Binary Access Read Shared As #f
                st = Space$(LOF(f))
                Get #f, , st
            Close #f
            
            Set GPC = New GDIPlusCandy
            Set sPic = GPC.DataToImage(st)
            Set GPC = Nothing
            
            If Not sPic Is Nothing Then
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                
                ' As ID3v2.0 and ID3v2.2 allow only JPEG and PNG images, do the necessary conversion for BMP and GIF images
                If bConvertImage Then
                    Set GPC = New GDIPlusCandy
                    st = GPC.ImageToData(sPic, ImagePNG)
                    Set sPic = GPC.DataToImage(st) ' Show the converted image
                    Set GPC = Nothing
                End If
                
                tMIMEType = DetermineImageType(st, ID3Revision)
                If sMIMEType <> tMIMEType And tMIMEType <> ImageUnsupported Then
                    sMIMEType = tMIMEType
                    cmbImageType.ListIndex = GetIndex(sMIMEType, ID3Revision)
                End If
                ArtAddProc sMIMEType, sPic, st
            End If
        End If
    'End If
End Sub

Private Sub TextProc(Ctl As Object, ByVal Description As String, CountControl As Label, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    If Index = Total Then
        If Ctl = "" Then
            If Index > 0 Then
                FrameBlank = True
                Col.Remove Index
                Index = Index - 1
                Total = Total - 1
                Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                NextControl.ToolTipText = ""
                If Index = 0 Then
                    Set DelControl.Picture = Buttons.ListImages(I_DELI).Picture
                    DelControl.ToolTipText = ""
                End If
            End If
        Else
            If FrameBlank Then
                FrameBlank = False
                Col.Add Ctl.Text
                Index = Index + 1
                Total = Total + 1
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            Else
                If Index > 0 Then SetItem Col, Index, Ctl.Text
            End If
        End If
    Else
        If Index > 0 Then SetItem Col, Index, Ctl.Text
    End If
    CountControl = CStr(Index) & "/" & CStr(Total)
End Sub

Private Sub ArtAddProc(ByVal MIMEType As String, ByVal Pic As StdPicture, ByVal Data As String)
    picArt.ToolTipText = S_APICTT
    imgArt.ToolTipText = S_APICTT
    imgArt.Visible = True
    Set imgArt.Picture = Nothing
    StretchImage Pic
    Set imgArt.Picture = Pic
    SetBG True
    lblBrowse.Visible = False
    
    If indAPIC = totAPIC Then
        If bAPICBlank Then
            bAPICBlank = False
            AddAPICItem MIMEType, cmbPictureType.ListIndex, Data
            indAPIC = indAPIC + 1
            totAPIC = totAPIC + 1
            Set nextArt.Picture = Buttons.ListImages(I_ADD).Picture
            nextArt.ToolTipText = S_ADD & S_APIC
            Set delArt.Picture = Buttons.ListImages(I_DEL).Picture
            delArt.ToolTipText = S_DEL & S_APIC
        Else
            If indAPIC > 0 Then
                SetItem cAPICData, indAPIC, Data
                SetItem cAPICIType, indAPIC, MIMEType
                SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
            End If
        End If
    Else
        If indAPIC > 0 Then
            SetItem cAPICData, indAPIC, Data
            SetItem cAPICIType, indAPIC, MIMEType
            SetItem cAPICType, indAPIC, cmbPictureType.ListIndex
        End If
    End If
    countArt = CStr(indAPIC) & "/" & CStr(totAPIC)
End Sub

Private Sub PrevProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean, Optional ByVal DeleteMode As Boolean = False)
    Dim bPic As Boolean, GDP As GDIPlusCandy, Pic As StdPicture
    bPic = (TypeName(Ctl) = "PictureBox")
    If (Index > 0 And DeleteMode) Or (Index > 1 And Not DeleteMode) Or (Index > 0 And Not DeleteMode And FrameBlank) Then
        If Index = Total Then
            Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
            DelControl.ToolTipText = S_DEL & Description
            If FrameBlank Then
IsBlank:
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
                FrameBlank = False
            Else
                Index = Index - 1
                If Index = 0 Then
                    GoTo IsBlank
                Else
                    Set NextControl.Picture = Buttons.ListImages(I_NEXT).Picture
                    NextControl.ToolTipText = S_NEXT & Description
                End If
            End If
        Else
            Index = Index - 1
        End If
        If Index <= 1 Then
            Set PrevControl.Picture = Buttons.ListImages(I_PREVI).Picture
            PrevControl.ToolTipText = ""
        End If
        If Index = 0 Then
            If bPic Then
                SetBG False
                Set imgArt.Picture = Nothing
                StretchImage imgArt.Picture
                imgArt.Visible = False
                cmbImageType.Enabled = False
                cmbPictureType.Enabled = False
                cmbImageType.ListIndex = 2 * (cmbImageType.ListIndex \ 4)
                cmbPictureType.ListIndex = 0
                lblBrowse.Visible = True
                Ctl.ToolTipText = ""
                imgArt.ToolTipText = ""
            Else
                Ctl = ""
            End If
            FrameBlank = True
        Else
            If bPic Then
                lblBrowse.Visible = False
                imgArt.Visible = True
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                Set GDP = New GDIPlusCandy
                Set Pic = GDP.DataToImage(Col(Index))
                Set GDP = Nothing
                Set imgArt.Picture = Nothing
                StretchImage Pic
                Set imgArt.Picture = Pic
                SetBG True
                cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                cmbPictureType.ListIndex = cAPICType(Index)
                Ctl.ToolTipText = S_APICTT
                imgArt.ToolTipText = S_APICTT
            Else
                Ctl = Col(Index)
            End If
        End If
        CountControl = CStr(Index) & "/" & CStr(Total)
    End If
End Sub

Private Sub NextProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim bPic As Boolean, bBlank As Boolean, GDP As GDIPlusCandy, Pic As StdPicture, lIType As Long, vType As Variant
    Dim bRefresh As Boolean
    Dim blFrameBlank As Boolean
    bPic = (TypeName(Ctl) = "PictureBox")
    If Index < Total Then
        bRefresh = True
        Index = Index + 1
        If Index = Total Then
            Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
            NextControl.ToolTipText = S_ADD & Description
        End If
    Else
        If bPic Then
            If imgArt.Visible Then
                bRefresh = True
                blFrameBlank = True
                bBlank = True
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            End If
        Else
            If Ctl <> "" Then
                bRefresh = True
                blFrameBlank = True
                Index = Index + 1
                Col.Add ""
                Total = Col.Count
                Set DelControl.Picture = Buttons.ListImages(I_DEL).Picture
                DelControl.ToolTipText = S_DEL & Description
            End If
        End If
    End If
    If bRefresh Then
        Set PrevControl.Picture = Buttons.ListImages(I_PREV).Picture
        PrevControl.ToolTipText = S_PREV & Description
        If bPic Then
            If bBlank Then
                cmbImageType.Enabled = False
                cmbPictureType.Enabled = False
                imgArt.Visible = False
                SetBG False
                Set imgArt.Picture = Nothing
                StretchImage imgArt.Picture
                cmbImageType.ListIndex = 2 * (cmbImageType.ListCount \ 4)
                cmbPictureType.ListIndex = 0
                lblBrowse.Visible = True
                Ctl.ToolTipText = ""
                imgArt.ToolTipText = ""
                Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                NextControl.ToolTipText = ""
            Else
                cmbImageType.Enabled = True
                cmbPictureType.Enabled = True
                imgArt.Visible = True
                Set GDP = New GDIPlusCandy
                Set Pic = GDP.DataToImage(Col(Index))
                Set GDP = Nothing
                Set imgArt.Picture = Nothing
                StretchImage Pic
                Set imgArt.Picture = Pic
                SetBG True
                cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                cmbPictureType.ListIndex = cAPICType(Index)
                lblBrowse.Visible = False
                Ctl.ToolTipText = S_APICTT
                imgArt.ToolTipText = S_APICTT
                Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                NextControl.ToolTipText = S_ADD & Description
            End If
        Else
            Ctl = Col(Index)
        End If
        If blFrameBlank Then FrameBlank = True
        CountControl = CStr(Index) & "/" & CStr(Total)
    End If
End Sub

Private Sub DelProc(Ctl As Object, ByVal Description As String, CountControl As Label, PrevControl As Image, NextControl As Image, DelControl As Image, Col As Collection, Index As Long, Total As Long, FrameBlank As Boolean)
    Dim bPic As Boolean, GDP As GDIPlusCandy, Pic As StdPicture
    bPic = (TypeName(Ctl) = "PictureBox")
    If Total > 0 Then
        If FrameBlank Then
            PrevProc Ctl, Description, CountControl, PrevControl, NextControl, DelControl, Col, Index, Total, FrameBlank, True
        Else
            If bPic Then
                'RemoveAPICItem Index
            Else
                Col.Remove Index
            End If
            Total = Col.Count
            If Index > Total Then Index = Total
            If Index = 0 Then
                If bPic Then
                    cmbImageType.Enabled = False
                    cmbPictureType.Enabled = False
                    SetBG False
                    Set imgArt.Picture = Nothing
                    StretchImage imgArt.Picture
                    cmbImageType.ListIndex = 2 * (cmbImageType.ListIndex \ 4)
                    cmbPictureType.ListIndex = 0
                    imgArt.Visible = False
                    lblBrowse.Visible = True
                    Ctl.ToolTipText = ""
                    imgArt.ToolTipText = ""
                Else
                    Ctl = ""
                End If
                FrameBlank = True
            Else
                If bPic Then
                    cmbImageType.Enabled = True
                    cmbPictureType.Enabled = True
                    imgArt.Visible = True
                    Set GDP = New GDIPlusCandy
                    Set Pic = GDP.DataToImage(Col(Index))
                    Set GDP = Nothing
                    Set imgArt.Picture = Nothing
                    StretchImage Pic
                    Set imgArt.Picture = Pic
                    SetBG True
                    cmbImageType.ListIndex = GetIndex(cAPICIType(Index), ID3Revision)
                    cmbPictureType.ListIndex = cAPICType(Index)
                    lblBrowse.Visible = False
                    Ctl.ToolTipText = S_APICTT
                    imgArt.ToolTipText = S_APICTT
                Else
                    Ctl = Col(Index)
                End If
            End If
            If Index = Total Then
                If Index = 0 Then
                    Set NextControl.Picture = Buttons.ListImages(I_ADDI).Picture
                    NextControl.ToolTipText = ""
                Else
                    Set NextControl.Picture = Buttons.ListImages(I_ADD).Picture
                    NextControl.ToolTipText = S_ADD & Description
                End If
            End If
            If Index <= 1 Then
                Set PrevControl.Picture = Buttons.ListImages(I_PREVI).Picture
                PrevControl.ToolTipText = ""
                If Total = 0 Then
                    Set DelControl.Picture = Buttons.ListImages(I_DELI).Picture
                    DelControl.ToolTipText = ""
                End If
            End If
            CountControl = CStr(Index) & "/" & CStr(Total)
        End If
    End If
End Sub



Private Function WithinBounds(ByVal Obj As Object, ByVal X As Single, ByVal Y As Single) As Boolean
    Dim oWidth As Single, oHeight As Single
    If TypeName(Obj) = "PictureBox" Then
        oWidth = Obj.ScaleWidth
        oHeight = Obj.ScaleHeight
    Else
        oWidth = Obj.Width
        oHeight = Obj.Height
    End If
    WithinBounds = (X >= 0 And X <= oWidth And Y >= 0 And Y <= oHeight)
End Function
