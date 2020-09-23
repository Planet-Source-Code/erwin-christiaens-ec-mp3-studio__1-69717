VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form FrmFileDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialog"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12210
   Icon            =   "FrmFileDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "move"
      Height          =   285
      Left            =   75
      TabIndex        =   14
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   270
      Left            =   150
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6375
      Left            =   60
      TabIndex        =   7
      Top             =   315
      Width           =   3255
      Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
         Height          =   5130
         Left            =   135
         TabIndex        =   10
         Top             =   255
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   9049
         BorderStyle     =   1
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Walk through subdirecotires"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   6000
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   6375
      Left            =   3300
      TabIndex        =   2
      Top             =   315
      Width           =   8895
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   7545
         TabIndex        =   13
         Top             =   5895
         Width           =   1230
      End
      Begin VB.CommandButton Cmdopen 
         Caption         =   "Open"
         Height          =   345
         Left            =   7515
         TabIndex        =   12
         Top             =   5505
         Width           =   1230
      End
      Begin MSComctlLib.ListView lv 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9128
         View            =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "path"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   5520
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "no of files"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   5880
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "size"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   6105
         Width           =   330
      End
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1020
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6810
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   5100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6810
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   9060
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":0F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":15C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":16D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":19F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":1D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":202A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":2346
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":2662
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":297E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":2C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":2FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":32D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":394E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":3C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":4546
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":476E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":4DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":5106
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFileDialog.frx":5262
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   4980
      TabIndex        =   9
      Top             =   7050
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "FrmFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const m_cSW_SHOW = 5

Private Enum m_e_FileAttributes
    
    Normal = 0
    ReadOnly = 1
    Hidden = 2
    System = 4
    volume = 8
    Directory = 16
    Archive = 32
    Alias = 64
    Compressed = 128
    
End Enum

Private m_lngIntIndex                  As Long
Private m_strOutputFile             As String
Private m_strPath                   As String
Private m_strWinPath                As String
Private m_strWinAppPath             As String
Private m_strWinSysPath             As String
Private m_astrStrSize()                 As String
Private m_astrStrAttr()                 As String
Private m_astrStrType()                 As String
Private m_astrStrText()                 As String
Private m_astrStrDateModified()         As String
Private m_FSO                       As New Scripting.FileSystemObject

Public CurrentDir As String
Sub SortColumn(Column As Integer)
lv.Sorted = True
lv.SortKey = Column
lv.SortOrder = lvwAscending
End Sub

Private Sub DoList()
  SortColumn (0) 'sortiranje_kolone (0)
  prekini_proces = False
  Call ListFiles
  
  If lv.ListItems.Count = 0 Then
    editirano = False
  Else
    editirano = True
  End If
End Sub

Sub ListFiles()
   Dim SearchPath As String, FindStr As String
    Dim filesize As Double
    Dim NumFiles As Long, NumDirs As Long
    Screen.MousePointer = vbHourglass
    
    
    lv.ListItems.Clear

    
    SearchPath = CurrentDir
    FindStr = "*.mp3"
    filesize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs, Check1.Value)
    Screen.MousePointer = vbDefault

Label2 = ""
Label3 = ""
Label4 = ""
Label1 = UCase(Dir1)
For X = 1 To lv.ListItems.Count
Label2 = "Number of MP3 files: " & X
Next X
List1.Clear
List2.Clear

Label5 = ""


Label3 = "Total size: " & veaf & " MB"
End Sub





Private Sub CmdClose_Click()
Form5.List1.Clear
Unload Me
End Sub

Private Sub Cmdopen_Click()
Dim X As Long
Form5.List1.Clear
List2.Clear
For X = 1 To lv.ListItems.Count
  If lv.ListItems(X).Checked = True Then
    'If lv.ListItems(x3).Checked = True Then
    Form5.List1.AddItem lv.ListItems.Item(X).SubItems(1)
    'Form5.List1.AddItem lv.ListItems.Item(X)
  End If
Next X
Unload Me
End Sub

Private Sub Command1_Click()
FrmFind.Show vbModal, Me
End Sub

Private Sub Command2_Click()
  Dim Folder As String
  Dim sExistingFolder As String
  
  If Right$(FolderTreeview1.SelectedFolder, 1) = "\" Then
    sExistingFolder = FolderTreeview1.SelectedFolder
  Else
    sExistingFolder = FolderTreeview1.SelectedFolder & "\"
  End If
  Folder = BrowseForFolder(hWnd, "Select a folder:", sExistingFolder)
  
  If Folder = "" Then
    MsgBox ("You must select path where you will move files!"), vbCritical
    
    Exit Sub
  End If
  If Folder <> "" Then
    On Error Resume Next
  If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
haha:
     For i = 1 To FrmFileDialog.lv.ListItems.Count
      If FrmFileDialog.lv.ListItems.Item(i).Checked = True Then
        FileCopy lv.ListItems.Item(i).SubItems(1), Folder & lv.ListItems.Item(i)
        Kill lv.ListItems.Item(i).SubItems(1)
  
          FrmFileDialog.lv.ListItems.Remove (i)
          GoTo haha:
          Exit For
      End If
    Next i
  End If
End Sub

Private Sub FolderTreeview1_FolderClick(p_Folder As CCRPFolderTV6.Folder, p_Location As CCRPFolderTV6.ftvHitTestConstants)
    ' Comments  :
    ' Parameters: p_Folder
    '             p_Location -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    Dim strPath             As String
    Dim pFolder             As Scripting.Folder
    Dim pFiles              As Scripting.Files
    Dim pFile               As Scripting.File
    Dim lngIndex            As Long
    
    
    'StatusBar1.Panels.Item(1).Text = p_Folder.DisplayName
    'StatusBar1.Panels.Item(2).Text = ""
    
    
    If (Right(p_Folder.FullPath, 1) = "}") Then
        strPath = vbNullString
    Else
        strPath = p_Folder.FullPath
    End If
    
    
    If strPath = vbNullString Then
        
        If p_Folder.Name = "Desktop" Then
            
            strPath = FolderTreeview1.GetSpecialFolderName(ftvDesktopDir)
            Set pFolder = m_FSO.GetFolder(strPath)
            Set pFiles = pFolder.Files
            If Right(strPath, 1) = "\" Then
                m_strPath = strPath
            Else
                m_strPath = strPath + "\"
            End If
            DoEvents
        CurrentDir = m_strPath
        DoList
            GoTo PROC_EXIT
        End If
        
        strPath = "c:\" & p_Folder.Name
        
        If strPath = "c:\Recycle Bin" Then
            strPath = "c:\recycled"
            
            Set pFolder = m_FSO.GetFolder(strPath)
            Set pFiles = pFolder.Files
            If Right(strPath, 1) = "\" Then
                m_strPath = strPath
            Else
                m_strPath = strPath + "\"
            End If
            DoEvents
        CurrentDir = m_strPath
        DoList
            GoTo PROC_EXIT
        Else
            'lv.Icons = Nothing
            'ListView1.SmallIcons = Nothing
            'ListView1.ListItems.Clear
            'ImageList1.ListImages.Clear
            'ImageList2.ListImages.Clear
            GoTo PROC_EXIT
        End If
        
    Else
        
        Set pFolder = m_FSO.GetFolder(strPath)
        Set pFiles = pFolder.Files
        If Right(strPath, 1) = "\" Then
            m_strPath = strPath
        Else
            m_strPath = strPath + "\"
        End If
        DoEvents
        CurrentDir = m_strPath
        DoList
        GoTo PROC_EXIT
    End If
    
    
    
    
    
PROC_EXIT:
    lv.SetFocus
    DoEvents
    Me.Refresh
    ' Clean Up
    Set pFolder = Nothing
    Set pFiles = Nothing
    Set pFile = Nothing
    Exit Sub
    
PROC_ERR:
    If err.Number = 5 Or err.Number = 76 Then
        Resume PROC_EXIT
    Else
        MsgBox err.Description
        Resume PROC_EXIT
    End If
    
End Sub

Private Sub FolderTreeview1_MouseUp(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
    ' Comments  :
    ' Parameters: p_intButton
    '             p_intShift
    '             p_sngX
    '             p_sngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    If p_intButton = vbRightButton Then
        
        PopupMenu mnuPopUp
        
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub Form_Load()
With lv
.View = lvwReport 'lvwList
.ColumnHeaders.Add , , "Filename"
.ColumnHeaders.Add , , "Path", 0 'lv.Width ' vbCenter
.ColumnHeaders.Add , , "Song Title" ', lv.Width * 0.2, vbCenter
.ColumnHeaders.Add , , "Artist"
.ColumnHeaders.Add , , "Album"
.ColumnHeaders.Add , , "Duration"
.ColumnHeaders.Add , , "Size (MB)"
End With

Dim strPath             As String
Dim lngIndex            As Long

'dblWidth = ListView1.Width
m_strWinAppPath = "C:\Documents and Settings\Eigenaar\Shared\"

FolderTreeview1.SelectedFolder = m_strWinAppPath
m_strPath = strPath


End Sub

Private Sub lv_DblClick()
X = lv.selectedItem.Index
FrmMp3TagInfo.filename = lv.ListItems.Item(X).SubItems(1)
FrmMp3TagInfo.Show vbModal
End Sub
