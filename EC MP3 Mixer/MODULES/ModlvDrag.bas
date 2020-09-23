Attribute VB_Name = "ModlvDrag"
Option Explicit

'************************************************************************
' Hi all, the following code is a sample of how you can work with
' ListViews. I love the 'Drag and Drop' functionality and would like to
' see someone come up with a nice and funky app' to make some real use of it.
' So make good use of the code and leave some notes on what you think, plus
' any suggestions etc.
'
' Cheers and regards,
'
' The GazMan November 2002
'
'************************************************************************

Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' ListView functions
Public Type LVHITTESTINFO
  pt As POINTAPI
  Flags As LVHITTESTINFO_flags
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long
#End If
End Type

Public Enum LVHITTESTINFO_flags
  LVHT_NOWHERE = &H1   ' in LVW client area, but not over item
  LVHT_ONITEMICON = &H2
  LVHT_ONITEMLABEL = &H4
  LVHT_ONITEMSTATEICON = &H8
  LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

  'Outside the LVW's client area
  LVHT_ABOVE = &H8
  LVHT_BELOW = &H10
  LVHT_TORIGHT = &H20
  LVHT_TOLEFT = &H40
End Enum

Public Const LVM_FIRST = &H1000
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_HITTEST = (LVM_FIRST + 18)


Public AccessApp            As String
Public sSQL                 As String


'************************************************************************
' The following two func's find the position of the drop...
' The GazMan November 2002
'************************************************************************

Public Function ListView_GetItemPosition(hwndLV As Long, i As Long, ppt As POINTAPI) As Boolean

  ListView_GetItemPosition = SendMessage(hwndLV, LVM_GETITEMPOSITION, ByVal i, ppt)
  
End Function

Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long

  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
  
End Function

