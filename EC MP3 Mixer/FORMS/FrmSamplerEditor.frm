VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSamplerEditor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   Picture         =   "FrmSamplerEditor.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   390
      Picture         =   "FrmSamplerEditor.frx":85242
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   26
      Top             =   6075
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   780
      Picture         =   "FrmSamplerEditor.frx":85824
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   6075
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   8
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   6135
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   390
      Picture         =   "FrmSamplerEditor.frx":85E06
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   23
      Top             =   5385
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   780
      Picture         =   "FrmSamplerEditor.frx":863E8
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   22
      Top             =   5385
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   7
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5445
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   390
      Picture         =   "FrmSamplerEditor.frx":869CA
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   20
      Top             =   4710
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   780
      Picture         =   "FrmSamplerEditor.frx":86FAC
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   19
      Top             =   4710
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   6
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4770
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   405
      Picture         =   "FrmSamplerEditor.frx":8758E
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   17
      Top             =   4005
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   795
      Picture         =   "FrmSamplerEditor.frx":87B70
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   4005
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   5
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4065
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   375
      Picture         =   "FrmSamplerEditor.frx":88152
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   14
      Top             =   3315
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   765
      Picture         =   "FrmSamplerEditor.frx":88734
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   13
      Top             =   3315
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3375
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   390
      Picture         =   "FrmSamplerEditor.frx":88D16
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   11
      Top             =   2655
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   780
      Picture         =   "FrmSamplerEditor.frx":892F8
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   10
      Top             =   2655
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   3
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2715
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   375
      Picture         =   "FrmSamplerEditor.frx":898DA
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   1935
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   765
      Picture         =   "FrmSamplerEditor.frx":89EBC
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   1935
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1995
      Width           =   3030
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   390
      Picture         =   "FrmSamplerEditor.frx":8A49E
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   1245
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   780
      Picture         =   "FrmSamplerEditor.frx":8AA80
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   1245
      Width           =   345
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   1
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1305
      Width           =   3030
   End
   Begin VB.TextBox TxtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   0
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   630
      Width           =   3030
   End
   Begin VB.PictureBox PicPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   795
      Picture         =   "FrmSamplerEditor.frx":8B062
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   570
      Width           =   345
   End
   Begin VB.PictureBox Picbrowse 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   405
      Picture         =   "FrmSamplerEditor.frx":8B644
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   570
      Width           =   345
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   5220
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer SampPlayer 
      Height          =   405
      Left            =   1320
      TabIndex        =   27
      Top             =   7620
      Width           =   2430
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
      _cx             =   4286
      _cy             =   714
   End
End
Attribute VB_Name = "FrmSamplerEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub


Private Sub Form_Load()

  TxtFilename(0).Tag = GetSetting(App.Title, "Samplersound", "Path1", App.Path & "\wav\ELECTRIC A.wav")
  TxtFilename(0).Text = Mid(TxtFilename(0).Tag, InStrRevVB5(TxtFilename(0).Tag, "\") + 1, Len(TxtFilename(0).Tag))
  TxtFilename(1).Tag = GetSetting(App.Title, "Samplersound", "Path2", App.Path & "\wav\PECHE A.wav")
  TxtFilename(1).Text = Mid(TxtFilename(1).Tag, InStrRevVB5(TxtFilename(1).Tag, "\") + 1, Len(TxtFilename(1).Tag))
  TxtFilename(2).Tag = GetSetting(App.Title, "Samplersound", "Path3", App.Path & "\wav\WAVE_AHAHBPM.wav")
  TxtFilename(2).Text = Mid(TxtFilename(2).Tag, InStrRevVB5(TxtFilename(2).Tag, "\") + 1, Len(TxtFilename(2).Tag))
  TxtFilename(3).Tag = GetSetting(App.Title, "Samplersound", "Path4", App.Path & "\wav\FX BONGA.wav")
  TxtFilename(3).Text = Mid(TxtFilename(3).Tag, InStrRevVB5(TxtFilename(3).Tag, "\") + 1, Len(TxtFilename(3).Tag))
  TxtFilename(4).Tag = GetSetting(App.Title, "Samplersound", "Path5", App.Path & "\wav\WOOSHE A.wav")
  TxtFilename(4).Text = Mid(TxtFilename(4).Tag, InStrRevVB5(TxtFilename(4).Tag, "\") + 1, Len(TxtFilename(4).Tag))
  TxtFilename(5).Tag = GetSetting(App.Title, "Samplersound", "Path6", App.Path & "\wav\WOOSHE B.wav")
  TxtFilename(5).Text = Mid(TxtFilename(5).Tag, InStrRevVB5(TxtFilename(5).Tag, "\") + 1, Len(TxtFilename(5).Tag))
  TxtFilename(6).Tag = GetSetting(App.Title, "Samplersound", "Path7", App.Path & "\wav\LOOP 140 ELECTRO 05.wav")
  TxtFilename(6).Text = Mid(TxtFilename(6).Tag, InStrRevVB5(TxtFilename(6).Tag, "\") + 1, Len(TxtFilename(6).Tag))
  TxtFilename(7).Tag = GetSetting(App.Title, "Samplersound", "Path8", App.Path & "\wav\BPM.wav")
  TxtFilename(7).Text = Mid(TxtFilename(7).Tag, InStrRevVB5(TxtFilename(7).Tag, "\") + 1, Len(TxtFilename(7).Tag))
  TxtFilename(8).Tag = GetSetting(App.Title, "Samplersound", "Path9", App.Path & "\wav\LOOP 185 ANALOG 01.wav")
  TxtFilename(8).Text = Mid(TxtFilename(8).Tag, InStrRevVB5(TxtFilename(8).Tag, "\") + 1, Len(TxtFilename(8).Tag))
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Samplersound", "Path1", TxtFilename(0).Tag
  SaveSetting App.Title, "Samplersound", "Path2", TxtFilename(1).Tag
  SaveSetting App.Title, "Samplersound", "Path3", TxtFilename(2).Tag
  SaveSetting App.Title, "Samplersound", "Path4", TxtFilename(3).Tag
  SaveSetting App.Title, "Samplersound", "Path5", TxtFilename(4).Tag
  SaveSetting App.Title, "Samplersound", "Path6", TxtFilename(5).Tag
  SaveSetting App.Title, "Samplersound", "Path7", TxtFilename(6).Tag
  SaveSetting App.Title, "Samplersound", "Path8", TxtFilename(7).Tag
  SaveSetting App.Title, "Samplersound", "Path9", TxtFilename(8).Tag
  Unload Me
End Sub

Private Sub Picbrowse_Click(Index As Integer)
  On Error GoTo Error
  Dialog.CancelError = True
  Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
  Dialog.ShowOpen
  
  
  'Visual basic 6 users may want to get rid of the module...since it is a feature
  'that is already on VB6 (InStrRev)
  TxtFilename(Index).Text = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
  TxtFilename(Index).Tag = Dialog.filename
  Exit Sub
  
Error:
  If err.Number <> 32755 Then ' Cancel was pressed?
  MsgBox "Error loading file - " & err.Number & " : " & err.Description
  Else
  End If
End Sub

Private Sub PicPlay_Click(Index As Integer)
SampPlayer.URL = TxtFilename(Index).Tag
SampPlayer.Controls.play
End Sub
