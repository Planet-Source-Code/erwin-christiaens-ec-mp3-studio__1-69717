Attribute VB_Name = "modITunesRetrieval"
Option Explicit

Public ReceivedXML As String
Public bRet As Boolean
Public bLaunchArtDL As Boolean

Public bConnect As Boolean
Public dDuration As Double
Public bItunes As Boolean

Public ArtHost As String
Public ArtPort As Long
Public ArtPath As String

Public Sub AnalyzeURL(ByVal URL As String, Host As String, Port As Long, Path As String)
    Dim i As Long
    Dim j As Long
    Dim sURL As String
    
    sURL = URL
    i = InStr(sURL, "://")
    If i > 0 Then sURL = Mid$(sURL, i + 3)
    
    i = InStr(sURL, "/")
    j = InStr(sURL, ":")
    
    If i > 0 Then
        If j < i And j > 0 Then
            Host = Left$(sURL, j - 1)
            Port = Mid(sURL, j + 1, i - j - 1)
        Else
            Host = Left$(sURL, i - 1)
            Port = 80
        End If
        Path = Mid$(sURL, i)
    Else
        If j > 0 Then
            Host = Left$(sURL, j - 1)
            Port = Mid(sURL, j + 1)
        Else
            Host = sURL
            Port = 80
        End If
        Path = "/"
    End If
End Sub

Public Function HTTPCode(ByVal sData As String, Optional Description As String) As Long
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Description = ""
    If Left$(sData, 5) = "HTTP/" Then
        i = InStr(sData, " ")
        j = InStr(i + 1, sData, " ")
        k = InStr(j + 1, sData, vbCrLf)
        
        If i > 0 And j > 0 Then
            HTTPCode = CLng(Mid$(sData, i + 1, j - i - 1))
            
            If k > 0 Then
                Description = Mid$(sData, j + 1, k - j - 1)
            End If
        End If
    End If
End Function

Public Function GetTagData(ByVal Data As String, ByVal TagName As String, ByVal vType As VbVarType) As String
    Dim i As Long
    Dim j As Long
    Dim Snip As String
    Dim sType As String
    
    i = InStr(Data, "<key>" & TagName & "</key>")
    If i > 0 Then
        Snip = Mid$(Data, i + Len(TagName) + 11)
        
        If vType = vbString Then
            sType = "string"
        ElseIf vType = vbInteger Then
            sType = "integer"
        End If
        
        If Left$(Snip, 2 + Len(sType)) = "<" & sType & ">" Then
            j = InStr(Len(sType) + 2, Snip, "</" & sType & ">")
            If j > 0 Then
                'GetTagData = ReplaceHTML(UTF8toANSI(Mid$(Snip, Len(sType) + 3, j - Len(sType) - 3)))
            End If
        End If
    End If
End Function

Public Function GenerateHexString(ByVal Length As Long) As String
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    For i = 1 To Length
        Randomize ' Make sure we don't get the "sameness"
        r = Int(Rnd * 16) ' Generate random value between 0 and 15  ===>  [Random Integer] = Int(Rnd * (Max - Min + 1)) + Min
        Randomize ' Make sure we randomized the seed very well
        
        Select Case r
            Case 0 To 9
                s = s & CStr(r)
            Case Else
                s = s & Chr$(55 + r)
        End Select
    Next
    
    GenerateHexString = s
End Function


