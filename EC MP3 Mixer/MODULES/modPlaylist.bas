Attribute VB_Name = "modPlaylist"

Public Function IsBlank(Var As Variant) As Boolean

    If Len(Trim(Var)) = 0 Then
        IsBlank = True
    Else
        IsBlank = False
    End If
    
End Function
Public Sub LoadM3U(sPath As String, lvPlaylist As ListView)
Dim sBuff As String, i As Long, M3UChk As String * 7
    'On Error Resume Next
    Dim TrackNr As Integer
    Dim ID3 As New clsID3
    'Dim sPath As String
    Dim d As String
    'Dim HourPart As String
    'Dim BlankWCOM As New MultiFrameData
    'Dim BlankWOAR As New MultiFrameData
    'Dim BlankAPIC As New MultiFrameData
    
    If FileExists(sPath) = False Then Exit Sub
    
    '// Check for M3U Header
    Open sPath For Binary As #1
    Get 1#, 1, M3UChk
        If M3UChk <> "#EXTM3U" Then Exit Sub
    Close #1
    
    DoEvents
    TrackNr = 0
    '// Adding procedure
    Open sPath For Input As #1
    Do While Not EOF(1)
        Line Input #1, sBuff
            If IsBlank(sBuff) Then GoTo 1
            If Mid(sBuff, 1, 1) = "#" Then GoTo 1
            
            
            With lvPlaylist
                If MousePointer = vbDefault Then
                    MousePointer = vbHourglass
                    DoEvents
                End If
                FrmLoading.Label1.Caption = FilterMedia(StripFileName(sBuff))
                FrmLoading.Refresh
                DoEvents
                .ListItems.Add Text:=sBuff
                ID3.Filename = sBuff
                TrackNr = TrackNr + 1
                With .ListItems(.ListItems.Count)
                    .SubItems(1) = TrackNr & "."
                    .SubItems(2) = ID3.Title
                    .SubItems(3) = ID3.Artist
                    .SubItems(4) = ID3.Album
                    .SubItems(5) = Form5.FormatGenre(ID3, ID3.GenreID, ID3.Genre)
                    .SubItems(6) = ID3.TrackNumber
                    .SubItems(7) = ID3.TracksTotal
                    .SubItems(8) = ID3.Year
                    .SubItems(9) = Form5.FormatTime(ID3.Length)
                    .SubItems(10) = Form5.FormatBitRate(ID3.BitRate, ID3.Encoding)
                    .SubItems(11) = ID3.Comments
                End With
            End With
      'Resort = True
      'SortLvwOnLong lvPlaylist, lvPlaylist.SortKey + 1
      'Resort = False

            
            'lstPlaylist.AddItem FilterMedia(StripFileName(sBuff))
            'lstPlaylist_Path.AddItem sBuff
1
    Loop
    Close #1
    DoEvents
    Unload FrmLoading
End Sub
Public Function FilterMedia(strTitle As String) As String
Dim strBuff As String, varArray As Variant

    '// Basicly removes the file extenstion
    
    varArray = Split(strTitle, ".")
    
    strBuff = varArray(UBound(varArray))
    
    FilterMedia = Replace(strTitle, "." & strBuff, "")
    
End Function
Public Function StripFileName(FilePath As String) As String
Dim Path As Variant

On Error GoTo 1

    Path = Split(FilePath, "\")
    StripFileName = Path(UBound(Path))
    
1
End Function

Public Sub SaveM3U(sPath As String, lstPlaylist_Path As ListView)
Dim i As Long
    Close #1
    Open sPath For Output As #1
        Print #1, "#EXTM3U" '// m3u header
        For i = 0 To lstPlaylist_Path.ListItems.Count - 1
            Print #1, lstPlaylist_Path.ListItems(i + 1).Text 'print the file's path
        Next
    Close #1
End Sub


Public Function FileExists(sFilename As String) As Boolean
    If IsBlank(sFilename) Then
        FileExists = False
        Exit Function
    End If
    
    If Len(Dir$(sFilename)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

