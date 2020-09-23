Attribute VB_Name = "modWMRetrieval"
Option Explicit

Private Const META_TAB = "    "

Private Const META_TAB_1 = META_TAB
Private Const META_TAB_2 = META_TAB_1 & META_TAB
Private Const META_TAB_3 = META_TAB_2 & META_TAB
Private Const META_TAB_4 = META_TAB_3 & META_TAB

Public Const ART_HOST = "services.windowsmedia.com"
Public Const ART_PATH = "/cover/"

Public dBitRate As Double
Public ArtParam As String

Private Function SplitIntoWords(ByVal Expression As String) As String()
    Dim s() As String
    Dim t As String
    Dim t_0 As String
    Dim X As String
    Dim i As Long
    Dim j As Long
    
    s = Split(Expression, " ")
    
    For i = LBound(s) To UBound(s)
        t = s(i)
        t_0 = ""
        For j = 1 To Len(t)
            X = Mid$(t, j, 1)
            Select Case Asc(X)
                Case 45, 47 To 57, 65 To 90, 97 To 122, &HC0 To &HD6, &HD8 To &HF6, &HF8 To &HFF: t_0 = t_0 & X
            End Select
        Next
        s(i) = t_0
    Next
    
    SplitIntoWords = s
End Function

Private Function WordTags(ByVal Tabs As String, Words() As String) As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim s As String
    Dim subsplit() As String
    Dim subsubsplit() As String
    
    For i = LBound(Words) To UBound(Words)
        subsplit = Split(Words(i), "-")
        For j = LBound(subsplit) To UBound(subsplit)
            subsubsplit = Split(subsplit(j), "/")
            For k = LBound(subsubsplit) To UBound(subsubsplit)
                If subsubsplit(k) <> "" Then
                    's = s & Tabs & "<word>" & ANSItoUTF8(subsubsplit(k)) & "</word>" & vbCrLf
                End If
            Next
        Next
    Next
    
    WordTags = s
End Function

Public Function WGetTagData(ByVal Data As String, ByVal TagName As String, Optional ByVal ConvertHTML As Boolean = False) As String
    Dim i As Long
    Dim j As Long
    Dim td As String
    
    i = InStr(Data, "<" & TagName & ">")
    If i > 0 Then
        j = InStr(i + Len(TagName) + 2, Data, "</" & TagName & ">")
        If j > 0 Then
            td = Mid$(Data, i + Len(TagName) + 2, j - i - Len(TagName) - 2)
            'If ConvertHTML Then td = ReplaceHTML(UTF8toANSI(td))
            WGetTagData = td
        End If
    End If
End Function

Public Function WGetTagData2(ByVal Data As String, ByVal TagName As String, ByVal TagName2 As String, ByVal Iterations As Long, ReachedEnd As Boolean, Optional ByVal ConvertHTML As Boolean = False)
    Dim i0 As Long
    Dim i As Long
    Dim j As Long
    Dim Snip As String
    Dim td As String
    
    For i0 = 1 To Iterations
        i = InStr(j + 1, Data, "<" & TagName & ">")
        If i = 0 Then
            ReachedEnd = True
            Exit Function
        End If
        
        j = InStr(i + Len(TagName) + 2, Data, "</" & TagName & ">")
        If j = 0 Then
            ReachedEnd = True
            Exit Function
        End If
    Next
    ReachedEnd = False
    
    Snip = Mid$(Data, i + Len(TagName) + 2, j - i - Len(TagName) - 2)
    i = InStr(Snip, "<" & TagName2 & ">")
    If i > 0 Then
        j = InStr(i + Len(TagName2) + 2, Snip, "</" & TagName2 & ">")
        If j > 0 Then
            td = Mid$(Snip, i + Len(TagName2) + 2, j - i - Len(TagName2) - 2)
            'If ConvertHTML Then td = ReplaceHTML(UTF8toANSI(td))
            WGetTagData2 = td
        End If
    End If
End Function

