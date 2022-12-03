Sub TotalOfGrouped()

    Dim PriDict As New Scripting.Dictionary
    Dim CompDict As New Scripting.Dictionary
    
    
    Dim PriTbl As ListObject
    Dim InputTbl As ListObject
    
    Dim ARPri As Variant
    Dim ARInput As Variant
    Dim FirstComp As Variant
    Dim SecondComp As Variant
    
    Dim Total As Long
    
    Dim Str As String
    
    
    
    Set PriTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set InputTbl = ThisWorkbook.Worksheets(2).ListObjects(1)
    
    
    ARPri = PriTbl.DataBodyRange
    ARInput = InputTbl.DataBodyRange
    
    For i = 1 To UBound(ARPri)
        
        PriDict(ARPri(i, 1)) = ARPri(i, 2)
        
    Next i
    
    Total = 0
    
    For i = 1 To UBound(ARInput)
        
        'MsgBox Len(ARInput(i, 1))
        
        FirstComp = Mid(ARInput(i, 1), 1, Len(ARInput(i, 1)))
        
       
       
        For j = 1 To Len(FirstComp)
        
            If i Mod 3 = 0 Then
            
            
                Str = Mid(FirstComp, j, 1)
                
                If InStr(ARInput(i - 1, 1), Str) > 0 And InStr(ARInput(i - 2, 1), Str) > 0 Then
                    
                  Total = Total + PriDict(Str)
                  Exit For
                  
                
                End If
            
            End If
            
            
        
        Next j
        
        
       
        
    Next i
    
    


    MsgBox "The Total is " & Total, vbOKOnly, "Total"
    


End Sub
