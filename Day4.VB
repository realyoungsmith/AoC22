Sub FullyContainedRange()

    
    Dim InTbl As Excel.ListObject
    
    Dim ArIn As Variant
    
    Dim Pair As Variant
    Dim Rng1 As Variant
    Dim Rng2 As Variant
    
    Dim Cnt As Long

    Set InTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = InTbl.DataBodyRange
    Cnt = 0
    
    For i = 1 To UBound(ArIn)
    
        Pair = Split(ArIn(i, 1), ",")
        Rng1 = Split(Pair(0), "-")
        Rng2 = Split(Pair(1), "-")
        
            
        If (CDbl(Rng1(0)) <= CDbl(Rng2(0)) And CDbl(Rng1(1)) >= CDbl(Rng2(1))) Or (CDbl(Rng2(0)) <= CDbl(Rng1(0)) And CDbl(Rng2(1)) >= CDbl(Rng1(1))) Then
        
            Cnt = Cnt + 1
            GoTo NextI
            
        End If
        
        
NextI:
        
    Next i
    
     MsgBox "Number of overlapping ranges is " & Cnt, vbOKOnly, "Number of fully overlapped ranges"
     

End Sub
Sub OverlappingRange()

    
    Dim InTbl As Excel.ListObject
    
    Dim ArIn As Variant
    
    Dim Pair As Variant
    Dim Rng1 As Variant
    Dim Rng2 As Variant
    
    Dim Cnt As Long

    Set InTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = InTbl.DataBodyRange
    Cnt = 0
    
    For i = 1 To UBound(ArIn)
    
        Pair = Split(ArIn(i, 1), ",")
        Rng1 = Split(Pair(0), "-")
        Rng2 = Split(Pair(1), "-")
        
            
        If (CDbl(Rng1(0)) <= CDbl(Rng2(0)) And CDbl(Rng1(1)) >= CDbl(Rng2(1))) Or (CDbl(Rng2(0)) <= CDbl(Rng1(0)) And CDbl(Rng2(1)) >= CDbl(Rng1(1))) Then
        
            Cnt = Cnt + 1
            GoTo NextI
            
        End If
        
        If (CDbl(Rng1(0)) >= CDbl(Rng2(0)) And CDbl(Rng1(0)) <= CDbl(Rng2(1))) Or (CDbl(Rng1(1)) >= CDbl(Rng2(0)) And CDbl(Rng1(1)) <= CDbl(Rng2(1))) Or (CDbl(Rng1(0)) <= CDbl(Rng2(1)) And CDbl(Rng1(0)) >= CDbl(Rng2(0))) Then
        
            Cnt = Cnt + 1
            GoTo NextI
            
        End If
NextI:
        
    Next i
    
    
     MsgBox "Number of overlapping ranges is " & Cnt, vbOKOnly, "Number of overlapped ranges"

End Sub
