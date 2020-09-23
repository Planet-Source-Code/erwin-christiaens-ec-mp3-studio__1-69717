VERSION 5.00
Begin VB.Form FrmSearch 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSearch.frx":0000
   ScaleHeight     =   2970
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   690
      TabIndex        =   8
      Top             =   1635
      Value           =   1  'Checked
      Width           =   165
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3885
      Picture         =   "FrmSearch.frx":2F57A
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   4
      Top             =   840
      Width           =   885
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3885
      Picture         =   "FrmSearch.frx":30534
      ScaleHeight     =   330
      ScaleWidth      =   885
      TabIndex        =   3
      Top             =   360
      Width           =   885
   End
   Begin VB.TextBox TxtAlbum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   900
      TabIndex        =   2
      Top             =   1170
      Width           =   2640
   End
   Begin VB.TextBox TxtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   915
      TabIndex        =   1
      Top             =   795
      Width           =   2640
   End
   Begin VB.ListBox lstFiles 
      Height          =   2205
      Left            =   345
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   5535
      Width           =   5535
   End
   Begin VB.TextBox TxtArtist 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   900
      TabIndex        =   0
      Top             =   405
      Width           =   2640
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   75
      TabIndex        =   5
      Top             =   5115
      Width           =   6000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wat...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   375
      Left            =   990
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Include subfolders"
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   885
      TabIndex        =   9
      Top             =   1650
      Width           =   1665
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ready."
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   165
      TabIndex        =   7
      Top             =   2640
      Width           =   510
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddFiles(ByRef FileList() As String, ByVal FileCnt As Long)
    Dim i As Long
    
    'Add the new files to the track list.
    If FileCnt = 0 Then Exit Sub
    ReDim Preserve RdFiles(nFiles + FileCnt - 1)
    
    Me.MousePointer = 11
    Me.Enabled = False
    
    For i = 0 To FileCnt - 1
        RdFiles(nFiles + i).SourceFile = FileList(i)
        lblStatus.Caption = "Reading files... " & Int(100 * ((i + 1) / FileCnt)) & "% done"
        DoEvents
    Next i
    nFiles = nFiles + FileCnt
    
    Me.Enabled = True
    Me.MousePointer = 0
    lblStatus.Caption = "Ready."
End Sub



Private Sub ShowFiles()
    Dim i As Long
    Dim ID3 As New clsID3
    Dim fname As String
    Label2.Visible = True
    Form5.List1.Clear
    For i = 0 To nFiles - 1
            fname = RdFiles(i).SourceFile
            ID3.filename = fname
            If InStr(LCase(ID3.Artist), LCase(TxtArtist)) And LCase(TxtArtist) <> "" Then
              Form5.List1.AddItem fname
            ElseIf InStr(LCase(ID3.Title), LCase(TxtTitle)) And LCase(TxtTitle) <> "" Then
              Form5.List1.AddItem fname
            ElseIf InStr(LCase(ID3.Album), LCase(TxtAlbum)) And LCase(TxtAlbum) <> "" Then
              Form5.List1.AddItem fname
            End If
            
            lblStatus.Caption = "Renaming files... " & Int(100 * ((i + 1) / nFiles)) & "% done"
            DoEvents
    Next i
    Label2.Visible = False
    If Form5.List1.ListCount < 0 Then
      MsgBox "Geen nummers gevonden."
    Else
    Unload Me
    End If
  
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Picture1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub Picture1_Click()
    Dim lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    Dim RdStrings() As String, nNewFiles As Long
    Dim sExistingFolder As String
    
    If Right$(Text1, 1) = "\" Then
      sExistingFolder = Form5.Text1
    Else
      sExistingFolder = Form5.Text1 & "\"
    End If
    
    sPath = BrowseForFolder(hWnd, "Select a folder:", sExistingFolder)
    Screen.MousePointer = vbHourglass
    'Search for mp3 files in the selected folder recursivly.
    lblStatus.Caption = "Searching for mp3 files..."
    If Check1.Value = vbChecked Then
        nNewFiles = FindFiles(sPath, ".mp3", RdStrings(), True)
    Else
        nNewFiles = FindFiles(sPath, ".mp3", RdStrings(), False)
    End If
    Screen.MousePointer = vbDefault
    If nNewFiles = 0 Then
        MsgBox "There were no mp3 files found in this folder.", vbInformation
    Else
        'Add the files.
        AddFiles RdStrings, nNewFiles
        'Show the files.
        Screen.MousePointer = vbHourglass
        ShowFiles
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Frmpics.PicOk(1).Image
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = Frmpics.PicOk(0).Image
End Sub

Private Sub Picture2_Click()
  If Form5.List1.ListCount < 0 Then
    Form5.List1.Clear
  End If
  Me.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Picture = Frmpics.Picannuleren(1).Image
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Picture = Frmpics.Picannuleren(0).Image
End Sub
