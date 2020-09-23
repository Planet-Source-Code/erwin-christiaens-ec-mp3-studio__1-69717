Attribute VB_Name = "mMain"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
         ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
         ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
         ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'APIs used to find files.
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long





Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1


Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const PROGRESS_CANCEL = 1
Public Const PROGRESS_CONTINUE = 0
Public Const PROGRESS_QUIET = 3
Public Const PROGRESS_STOP = 2
Public Const COPY_FILE_FAIL_IF_EXISTS = &H1
Public Const COPY_FILE_RESTARTABLE = &H2

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Const GENERIC_WRITE = &H40000000
Const GENERIC_READ = &H80000000
Const OPEN_EXISTING = 3
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FO_DELETE = &H3
Const FOF_NOCONFIRMATION = &H10

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Const ICC_BAR_CLASSES As Long = &H4
Private Const ICC_LISTVIEW_CLASSES As Long = &H1
Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const ICC_TAB_CLASSES As Long = &H8
Private Const ICC_UPDOWN_CLASS As Long = &H10
Private Const ICC_USEREX_CLASSES As Long = &H200
Private Const ICC_WIN95_CLASSES As Long = &HFF&

Private Type InitCommonControlsEx
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type

Public Enum MP3SourceEnum
    SOURCE_IDV1
    SOURCE_IDV2
    SOURCE_FILENAME
    SOURCE_USERENTRY
End Enum

Public Type MP3File
    SourceFile As String
    'SourceType As MP3SourceEnum
    
    'FileInterpretItems() As String
   ' FileInterpretItemCnt As Long
   ' FileInterpretArtist As Long
   ' FileInterpretTitle As Long
      
    'HasIDv1 As Boolean
    'HasIDv2 As Boolean
    'IDv1 As ID3Tag
    'IDv2 As ID3Tag
    'FileTag As ID3Tag
    'UserTag As ID3Tag
End Type
Global Choosing As Boolean


Global OrigPosition
Global OrigX
Global OrigY
'Vars for position of mouse
Global PosX
Global PosY
'Vars used instead of me.left and me.top(saves time and space)
Global x
Global y
Public RdFiles() As MP3File, nFiles As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef TLPINITCOMMONCONTROLSEX As InitCommonControlsEx) As Long
Public LvLst As ListItem


Public Sub Main()
    On Error Resume Next
    Dim iccex As InitCommonControlsEx
    With iccex
        .dwSize = Len(iccex)
        .dwICC = ICC_BAR_CLASSES Or ICC_LISTVIEW_CLASSES Or ICC_PROGRESS_CLASS Or ICC_STANDARD_CLASSES Or ICC_TAB_CLASSES Or ICC_UPDOWN_CLASS Or ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES
    End With
    InitCommonControlsEx iccex
    On Error GoTo 0
    Form5.Show
End Sub



Public Function InStrRevVB5(ByVal StringToCheck As String, ByVal StringToMatch As String, Optional ByVal StartAt As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
 
Dim lPos        As Long
Dim lSavePos    As Long
 
    ' -1 means search entire string. A positive number
    ' means search only up to that position from the left.
    If StartAt = -1 Then StartAt = Len(StringToCheck)
    
    ' Find the last instance of StringToMatch within StringToCheck.
    lPos = InStr(1, StringToCheck, StringToMatch, Compare)
    While lPos > 0 And lPos < StartAt
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringToCheck, StringToMatch, Compare)
    Wend
    
    InStrRevVB5 = lSavePos
        
End Function

Public Function BasePath(ByVal fname As String, Optional delim As String = "\", Optional keeplast As Boolean = True) As String
    Dim outstr As String
    Dim llen As Long
    llen = InStrRevVB5(fname, delim)


    If (Not keeplast) Then
        llen = llen - 1
    End If


    If (llen > 0) Then
        BasePath = Mid(fname, 1, llen)
    Else
        BasePath = fname
    End If
End Function

Public Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Long, DirCount As Long, subdir As Boolean)


    Dim ID3 As New clsID3
    

Dim filename As String                          ' Walking filename variable...
Dim DirName As String                           ' SubDirectory Name
Dim dirNames() As String                        ' Buffer for directory name entries
Dim nDir As Long                                    ' Number of directories in this path
Dim i As Long                                         ' For-loop counter...
Dim hSearch As Long                              ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Long

If Right(Path, 1) <> "\" Then Path = Path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(Path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DoEvents
DirName = StripNulls(WFD.cFileName)
'Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
'Check for directory with bitwise comparison.
If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD)                   'Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If
' Walk through this directory and sum file sizes.
hSearch = FindFirstFile(Path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
filename = StripNulls(WFD.cFileName)
'FrmFileDialog.Label5 = "Searching..... " & path


'naziv_pjesme = GetTag(path & filename)     'Za ispisivanje TAGOVA iz MP3 fajla
'naziv_pjesmev2 = GetTagID3v2(path & filename)

'art = nas
'album1 = alb
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If (filename <> ".") And (filename <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
If FrmFileDialog.lv.View = lvwReport Then
'/////// Uèitavanje u ListView /////////////////////////////////////////
Set LvLst = FrmFileDialog.lv.ListItems.Add(, , filename)
LvLst.SubItems(1) = Path & filename
With LvLst
    ID3.filename = Path & filename
    
    'With .ListItems(.ListItems.Count)
        '.SubItems(1) = TrackNr & "."
        .SubItems(2) = ID3.Title
        .SubItems(3) = ID3.Artist
        .SubItems(4) = ID3.Album
        '.SubItems(5) = Form5.FormatGenre(ID3, ID3.GenreID, ID3.Genre)
        '.SubItems(6) = ID3.TrackNumber
        '.SubItems(7) = ID3.TracksTotal
        '.SubItems(8) = ID3.Year
        .SubItems(5) = Form5.FormatTime(ID3.Length)
        .SubItems(6) = Form5.FormatBitRate(ID3.BitRate, ID3.Encoding)
        '.SubItems(11) = ID3.Comments
    'End With
End With
End If
'////////////////////////////////////////////////////////////////////////////////

End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)

End If
'If there are sub-directories...
If subdir = False Then
Exit Function
Else
If nDir > 0 Then
'Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount, subdir)
Next i
End If
End If
End Function

Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function


Public Function SetOpstartDir(ByVal NewValue As String)
    Dim sValue As String
    SaveSetting App.Title, "Opstart", "dir", NewValue
End Function
Public Function GetOpstartDir() As String
    Dim sValue As String
    'default
    GetOpstartDir = "c:\"
    sValue = GetSetting(App.Title, "Opstart", "dir", "")
    GetOpstartDir = sValue
End Function


Public Function FindFiles(ByVal Path As String, ByVal SearchStr As String, ByRef outFiles() As String, ByVal SubDirs As Boolean, Optional ByRef nFiles As Long = 0) As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    
    Path = NormalizeDir(Path)
    
    'Walk through this directory and get matching files.
    hSearch = FindFirstFile(Path & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                'Valid item.
                If SubDirs And (GetFileAttributes(Path & CurItem) And FILE_ATTRIBUTE_DIRECTORY) Then
                    'Item is a sub-directory, read it recursivly.
                    FindFiles Path & CurItem, SearchStr, outFiles(), True, nFiles
                ElseIf InStr(1, Path & CurItem, SearchStr, vbTextCompare) Then
                    'Item is a file which we're searching for.
                    ReDim Preserve outFiles(nFiles)
                    outFiles(nFiles) = Path & CurItem
                    nFiles = nFiles + 1
                End If
            End If
            'Get next item
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
    End If
    'Return the number of files in this directory (well, is also stored in the ByRef parameter nFiles).
    FindFiles = nFiles
End Function
'***File functions***

Public Function NormalizeDir(ByVal sDir As String) As String
    If Not Right$(sDir, 1) = "\" Then sDir = sDir & "\"
    NormalizeDir = sDir
End Function

Public Function GetDir(ByVal sPath As String) As String
    GetDir = NormalizeDir(Left$(sPath, Len(sPath) - InStr(1, StrReverse(sPath), "\") + 1))
End Function

Public Function GetFile(ByVal sPath As String) As String
    If InStr(sPath, "\") = 0 Then
        GetFile = sPath
    Else
        GetFile = Right$(sPath, InStr(1, StrReverse(sPath), "\") - 1)
    End If
End Function

Public Function GetFileWOExt(ByVal sPath As String) As String
    sPath = GetFile(sPath)
    If InStr(sPath, ".") = 0 Then
        GetFileWOExt = sPath
    Else
        GetFileWOExt = Left$(sPath, Len(sPath) - InStr(1, StrReverse(sPath), "."))
    End If
End Function
