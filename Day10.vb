Sub Part1()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    
    Dim InstrDict As New Scripting.Dictionary
    Dim CycleDict As New Scripting.Dictionary
    
    Dim SplitIn As Variant
      
    Dim Cycle As Long
    Dim X As Long
    Dim SignalSums As Long
    
    Dim Op As String
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = TblIn.DataBodyRange
    
    InstrDict("addx") = 2
    InstrDict("noop") = 1
    
    
    Cycle = 1
    X = 1
    For i = 1 To UBound(ArIn)
        
        If ArIn(i, 1) = "noop" Then
            Op = ArIn(i, 1)
        Else
            SplitIn = Split(ArIn(i, 1), " ")
            Op = SplitIn(0)
        End If
        
        Select Case Op
        
        Case "noop"
            Cycle = Cycle + 1
            
        Case "addx"
            Cycle = Cycle + 1
            CycleDict(CStr(Cycle)) = X
            
            Cycle = Cycle + 1
            SplitIn = Split(ArIn(i, 1), " ")
            
            X = X + CLng(SplitIn(1))
            
        End Select
        
        CycleDict(CStr(Cycle)) = X
        
    Next i
    
    SignalSum = 0
    For i = 20 To 220 Step 40
    
        SignalSum = SignalSum + (i * CycleDict(CStr(i)))
        
    
    Next i
    
    MsgBox SignalSum, vbOKOnly, "Signal Sum"
    
    
End Sub
Sub Part2()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    Dim TblOut As ListObject
    
    Dim InstrDict As New Scripting.Dictionary
    Dim CycleDict As New Scripting.Dictionary
    
    Dim SplitIn As Variant
      
    Dim Cycle As Long
    Dim X As Long
    Dim SignalSums As Long
    
    Dim Op As String
    
    Dim CRT As Variant
        
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set TblOut = ThisWorkbook.Worksheets(2).ListObjects(1)
    
    ArIn = TblIn.DataBodyRange
    
    InstrDict("addx") = 2
    InstrDict("noop") = 1
    
    ReDim CRT(0 To 5, 0 To 39)
    
    
    
    Cycle = 1
    X = 1
    For i = 1 To UBound(ArIn)
        
        If ArIn(i, 1) = "noop" Then
            Op = ArIn(i, 1)
        Else
            SplitIn = Split(ArIn(i, 1), " ")
            Op = SplitIn(0)
        End If
        
        Select Case Op
        
        Case "noop"
            
            If ((Cycle - 1) Mod 40) = X Or ((Cycle - 1) Mod 40) = X - 1 Or ((Cycle - 1) Mod 40) = X + 1 Then
            
                CRT(((Cycle - 1) \ 40), ((Cycle - 1) Mod 40)) = "#"
            
            End If
            
            
            Cycle = Cycle + 1
            
        Case "addx"
            If ((Cycle - 1) Mod 40) = X Or ((Cycle - 1) Mod 40) = X - 1 Or ((Cycle - 1) Mod 40) = X + 1 Then
            
                CRT(((Cycle - 1) \ 40), ((Cycle - 1) Mod 40)) = "#"
            
            End If
            Cycle = Cycle + 1
            TblOut.HeaderRowRange.Offset(1, 0).Resize(UBound(CRT, 1) + 1, UBound(CRT, 2) + 1) = CRT
            
            If ((Cycle - 1) Mod 40) = X Or ((Cycle - 1) Mod 40) = X - 1 Or ((Cycle - 1) Mod 40) = X + 1 Then
            
                CRT(((Cycle - 1) \ 40), ((Cycle - 1) Mod 40)) = "#"
            
            End If
            Cycle = Cycle + 1
            TblOut.HeaderRowRange.Offset(1, 0).Resize(UBound(CRT, 1) + 1, UBound(CRT, 2) + 1) = CRT
            SplitIn = Split(ArIn(i, 1), " ")
            
            X = X + CLng(SplitIn(1))
            
        End Select
        
        CycleDict(X) = X
        
    Next i
    
    

    
    

End Sub

