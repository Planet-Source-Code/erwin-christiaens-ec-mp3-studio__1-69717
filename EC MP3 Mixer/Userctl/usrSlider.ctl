VERSION 5.00
Begin VB.UserControl MSSlider 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ScaleHeight     =   345
   ScaleWidth      =   4080
   Begin VB.PictureBox picGripper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   108
      Picture         =   "usrSlider.ctx":0000
      ScaleHeight     =   165
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   60
      Width           =   240
   End
   Begin VB.PictureBox picGrove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   120
      ScaleHeight     =   120
      ScaleWidth      =   3855
      TabIndex        =   1
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "MSSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Event ValueHasChanged()

Private lColArray() As Long

Private lMin As Long
Private lMax As Long
Private lIndex As Long

Private lRange As Long
Private lStart As Long
Private lEnd As Long

Private lWidth As Long
Private lHeight As Long

Private lInitColor As Long
Private lFinalColor As Long

Private dblPercent As Double
Private bColorShift As Boolean

Property Let min(lMinimun As Long)
    lMin = lMinimun
End Property

Property Get min() As Long
    min = lMin
End Property

Property Let Max(lMaximum As Long)
    lMax = lMaximum
End Property

Property Get Max() As Long
    Max = lMax
End Property

Property Let Value(lCurrentPos As Long)
    Dim strPost As String
    
    lIndex = lCurrentPos
    dblPercent = (lIndex - lMin) / (lMax - lMin)
    SetPosFromValue
    RaiseEvent ValueHasChanged
End Property

Property Get Value() As Long
    Value = lIndex
End Property

Public Property Let LowColor(lS As Long)
    lInitColor = lS
End Property

Property Let HiColor(lE As Long)
    lFinalColor = lE
End Property

Property Let GripperPic(strFilename As String)
    LoadGripperPic (strFilename)
End Property

Property Let ColorShift(bShift As Boolean)
    bColorShift = bShift
End Property

Private Sub picGripper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents

    If Button = vbLeftButton Then
        Dim ax As Long
        Dim bx As Long

        ax = picGripper.Width / 2
        
        If (picGripper.Left + X - ax) >= 0 And ((picGripper.Left + picGripper.Width) + X - ax) <= (picGrove.Width) Then
            picGripper.Left = picGripper.Left + X - ax
        ElseIf (picGripper.Left + X - ax) < 0 Then
            picGripper.Left = 0
        ElseIf ((picGripper.Left + picGripper.Width) + X - ax) > (picGrove.Width) Then
            picGripper.Left = (picGrove.Width) - picGripper.Width
        End If
    
        SetValueFromPos
        SetColorFromPos
        RaiseEvent ValueHasChanged
        
    End If
    
End Sub

Private Function SetValueFromPos()
    Dim lDiff As Long
    dblPercent = ((picGripper.Left + picGripper.Width / 2) - lStart) / (lRange)
    lDiff = CInt(dblPercent * (lMax - lMin))
    
    lIndex = lDiff + lMin
    
End Function

Private Function SetColorFromPos()
    On Error Resume Next
    If bColorShift = True Then
        dblPercent = ((picGripper.Left + picGripper.Width / 2) - lStart) / (lRange)
        BlendColors lInitColor, lFinalColor, 100, lColArray
        picGrove.BackColor = lColArray(CInt(dblPercent * 100))
    End If
    
End Function

Private Function SetPosFromValue()
    dblPercent = ((lIndex) - lMin) / (lMax - lMin)
    picGripper.Left = ((dblPercent * lRange) + lStart) - (picGripper.Width / 2)

End Function

Private Sub picGripper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetValueFromPos
    SetColorFromPos
    RaiseEvent ValueHasChanged
End Sub

Private Sub picGrove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call picGripper_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    lMin = 1
    lMax = 100
    lIndex = 0
    picGripper.Left = 0
    picGrove.BackColor = &H8000000F
    

End Sub

Private Sub UserControl_Resize()
    picGrove.Left = 5
    picGrove.Top = 60
    picGrove.Height = UserControl.Height - 120
    picGrove.Width = UserControl.Width - 5
    
    picGripper.Left = picGrove.Left
    picGripper.ZOrder (0)
    
    If picGripper.Height > picGrove.Height Then
        picGripper.Top = picGrove.Top - (Abs(picGrove.Height - picGripper.Height) / 2)
    Else
        picGripper.Top = picGrove.Top + (Abs(picGrove.Height - picGripper.Height) / 2)
    End If

    
    lStart = (picGripper.Width / 2)
    lEnd = (picGrove.Width) - (picGripper.Width / 2)
    lRange = lEnd - lStart
    
    SetValueFromPos
    SetColorFromPos
 
End Sub

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    
    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

'Creates an array of colors blending from 'Color1 to Color2 in lSteps number of steps.
'Returns the count and fills the laRetColors() array.

Dim lIdx    As Long
Dim lRed    As Long
Dim lGrn    As Long
Dim lBlu    As Long
Dim fRedStp As Single
Dim fGrnStp As Single
Dim fBluStp As Single

    'Stop possible error
    If lSteps < 2 Then lSteps = 2
    
    'Extract Red, Blue and Green values from the start and end colors.
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    
    'Find the amount of change for each color element per color change.
    fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
    'Create the colors
    ReDim laRetColors(0 To lSteps)
    laRetColors(0) = lColor1            'First Color
    laRetColors(lSteps) = lColor2   'Last Color
    For lIdx = 1 To lSteps - 1          'All Colors between
        laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
            (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
            (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
    Next lIdx
    
    'Return number of colors in array
    BlendColors = lSteps

End Function

Public Function ColInit()
    SetColorFromPos
End Function

Private Function LoadGripperPic(strFilename)
    picGripper.Picture = LoadPicture(strFilename)
    UserControl_Resize
End Function
