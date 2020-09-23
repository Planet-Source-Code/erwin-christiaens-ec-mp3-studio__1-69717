Attribute VB_Name = "modSizeLimitAndWheel"
Option Explicit

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOVING As Long = &H216
Private Const WM_SIZE As Long = &H5
Private Const WM_SIZING As Long = &H214

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Public XPos As Long
Public YPos As Long

Public XSize As Long
Public YSize As Long

Private Declare Function DefWindowProc Lib "user32" Alias _
   "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias _
   "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
    ByVal cbCopy As Long)

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetComboBoxInfo Lib "user32" _
  (ByVal hwndCombo As Long, _
   CBInfo As COMBOBOXINFO) As Long



