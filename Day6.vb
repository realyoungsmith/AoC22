Sub FindPacketStart()
    
    Dim HoldDict As New Scripting.Dictionary
    
    
    Dim InTbl As ListObject
    
    Dim ArIn As Variant
    
    Dim InputStr As String
    
    Dim ChkString As String
    
    Dim CharString As String
    
    Dim Location As Long
    
    
    
    
    Set InTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = InTbl.DataBodyRange
    
    InputStr = ArIn
    
    
    Answer = 0
    For i = 1 To Len(InputStr)
        
        ChkString = Mid(InputStr, i, 14)
        
        For j = 1 To 14
            
            If Not HoldDict.Exists(Mid(ChkString, j, 1)) Then
            
                HoldDict(Mid(ChkString, j, 1)) = ""
                
                If j = 14 Then
                
                    Location = i + 13
                    GoTo Answer
                End If
                
                
           Else
           
           GoTo NextI
           
           
           End If
           
           
            
            
        
        Next j
        
        
        
NextI:
        
        HoldDict.RemoveAll
        
    Next i
Answer:
    
        MsgBox "Answer is " & Location, vbOKOnly, "Answer"
        
End Sub
