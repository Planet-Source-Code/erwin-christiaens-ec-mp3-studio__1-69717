Attribute VB_Name = "Module1"

'functions
Public Enum FindOptions
    PartOfWord = 0
    MatchCase = 1
    WholeWordOnly = 3
End Enum



Public Function FindLVItem(ByRef vLV As ListView, sCriteria As String, Optional iOption As FindOptions = 0, Optional MultiSelect As Boolean = False, Optional InverseSelection As Boolean = False, Optional FindNext As Boolean = False)

    Dim i As Integer
    Dim isFound As Boolean
    Dim li As Integer
    Dim StartPos As Integer
    
'On Error GoTo eh
    
    If vLV.ListItems.Count < 1 Then Exit Function

    If FindNext = True And vLV.selectedItem.Index < vLV.ListItems.Count Then
        For li = 1 To vLV.selectedItem.Index
            vLV.ListItems(li).Selected = False
        Next
        StartPos = vLV.selectedItem.Index + 1
    Else
        For li = 1 To vLV.ListItems.Count
            vLV.ListItems(li).Selected = False
        Next
        StartPos = 1
    End If
    
    'set flag to default
    isFound = False
    
    For li = StartPos To vLV.ListItems.Count
        
        Select Case iOption
            
            Case FindOptions.PartOfWord  'normal

                If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                        
                    isFound = True

                Else

                    'check subitems
                    For i = 1 To 3 'vLV.ListItems(li).ListSubItems.Count
                        If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                            Debug.Print vLV.ListItems(li).ListSubItems(1)
                            isFound = True
                            Exit For
                        
                        End If
                    Next
                                        
                End If
                
            Case FindOptions.MatchCase  'match case
            
            Case FindOptions.WholeWordOnly  ' whole word only
                
            
        End Select
        
        
        
        
        If isFound Then
            
            vLV.ListItems(li).Selected = CBool(True - InverseSelection)
            vLV.ListItems(li).EnsureVisible
            
            If Not MultiSelect Then Exit For
        
        Else
            vLV.ListItems(li).Selected = CBool(False - InverseSelection)
        End If
        
    Next
    
    If FindNext = True And isFound = False And StartPos > 1 Then
        
        For li = 1 To StartPos
            
            Select Case iOption
                
                Case FindOptions.PartOfWord  'normal
    
                    If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                            
                        isFound = True
    
                    Else
    
                        'check subitems
                        For i = 1 To vLV.ListItems(li).ListSubItems.Count
                            If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                                
                                isFound = True
                                Exit For
                            
                            End If
                        Next
                                            
                    End If
                    
                Case FindOptions.MatchCase  'match case
                
                Case FindOptions.WholeWordOnly  ' whole word only
                    
                
            End Select
            
            
            
            
            If isFound Then
                
                vLV.ListItems(li).Selected = CBool(True - InverseSelection)
                vLV.ListItems(li).EnsureVisible
                
                If Not MultiSelect Then Exit For
            
            Else
                vLV.ListItems(li).Selected = CBool(False - InverseSelection)
            End If
            
        Next
    End If
'On Error Resume Next
Exit Function
eh:
    MsgBox err.Description
    Resume Next
End Function

