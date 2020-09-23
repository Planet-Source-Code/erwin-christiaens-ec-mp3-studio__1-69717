VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFind 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoeken"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Find inside 'Song Title"""
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   2595
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Find inside 'Author'"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2595
      MaskColor       =   &H80000006&
      TabIndex        =   8
      Top             =   1425
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Find"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4995
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   5415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   5895
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Find whole word only"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000006&
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   6135
         Begin MSComctlLib.ListView lv 
            Height          =   3015
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   5318
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "FrmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1 = "" Then
Text1.SetFocus
Exit Sub
End If

lv.ListItems.Clear

If Check1.Value = 1 Then
    For X = 1 To FrmFileDialog.lv.ListItems.Count
        If FrmFileDialog.lv.ListItems(X).Text = Text1 Then
        Set bleh = lv.ListItems.Add(, , FrmFileDialog.lv.ListItems(X).Text)
        bleh.SubItems(1) = FrmFileDialog.lv.ListItems(X).SubItems(1)
        bleh.SubItems(2) = FrmFileDialog.lv.ListItems(X).SubItems(2)
        bleh.SubItems(3) = FrmFileDialog.lv.ListItems(X).SubItems(3)
        bleh.SubItems(4) = FrmFileDialog.lv.ListItems(X).SubItems(4)
        bleh.SubItems(5) = FrmFileDialog.lv.ListItems(X).SubItems(5)
        End If
    Next X
Else
    For X = 1 To FrmFileDialog.lv.ListItems.Count
        If InStr(FrmFileDialog.lv.ListItems(X).Text, Text1) > 0 Then
        ffs = Mid$(FrmFileDialog.lv.ListItems(X).Text, 1, Len(FrmFileDialog.lv.ListItems(X).Text))
        Set bleh = lv.ListItems.Add(, , ffs)
        bleh.SubItems(1) = FrmFileDialog.lv.ListItems(X).SubItems(1)
        bleh.SubItems(2) = FrmFileDialog.lv.ListItems(X).SubItems(2)
        bleh.SubItems(3) = FrmFileDialog.lv.ListItems(X).SubItems(3)
        bleh.SubItems(4) = FrmFileDialog.lv.ListItems(X).SubItems(4)
        bleh.SubItems(5) = FrmFileDialog.lv.ListItems(X).SubItems(5)
        End If
    Next X
End If

If Option1.Value = 1 And Check1.Value = 1 Then
    For X = 1 To FrmFileDialog.lv.ListItems.Count
        If FrmFileDialog.lv.ListItems(X).SubItems(2) = Text1 Then
        Set bleh = lv.ListItems.Add(, , FrmFileDialog.lv.ListItems(X).Text)
        bleh.SubItems(1) = FrmFileDialog.lv.ListItems(X).SubItems(1)
        bleh.SubItems(2) = FrmFileDialog.lv.ListItems(X).SubItems(2)
        bleh.SubItems(3) = FrmFileDialog.lv.ListItems(X).SubItems(3)
        bleh.SubItems(4) = FrmFileDialog.lv.ListItems(X).SubItems(4)
        bleh.SubItems(5) = FrmFileDialog.lv.ListItems(X).SubItems(5)
        End If
    Next X
Else
    For X = 1 To FrmFileDialog.lv.ListItems.Count
        If InStr(FrmFileDialog.lv.ListItems(X).SubItems(3), Text1) > 0 Then
        ffs = Mid$(FrmFileDialog.lv.ListItems(X).SubItems(3), 1, Len(FrmFileDialog.lv.ListItems(X).SubItems(3)))
        Set bleh = lv.ListItems.Add(, , ffs)
        bleh.SubItems(1) = FrmFileDialog.lv.ListItems(X).SubItems(1)
        bleh.SubItems(2) = FrmFileDialog.lv.ListItems(X).SubItems(2)
        bleh.SubItems(3) = FrmFileDialog.lv.ListItems(X).SubItems(3)
        bleh.SubItems(4) = FrmFileDialog.lv.ListItems(X).SubItems(4)
        bleh.SubItems(5) = FrmFileDialog.lv.ListItems(X).SubItems(5)
        End If
    Next X
End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
With lv
.View = lvwReport
.ColumnHeaders.Add , , "Filename"
.ColumnHeaders.Add , , "Path"
.ColumnHeaders.Add , , "Song Title"
.ColumnHeaders.Add , , "Author"
.ColumnHeaders.Add , , "Album"
.ColumnHeaders.Add , , "Duration"
End With
End Sub

