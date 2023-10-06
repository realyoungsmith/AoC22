Sub MoveCrates()
    
    Dim Stacks As New Scripting.Dictionary
    
    Dim InTbl As Excel.ListObject
    Dim StackTbl As Excel.ListObject
    
    Dim ArIn As Variant
    Dim ArStack As Variant
    Dim SplitIn As Variant
    
    Dim Qty As Long
    Dim FromStack As Long
    Dim ToStack As Long
    
    Dim StackTops As String
    
    

    Set InTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set StackTbl = ThisWorkbook.Worksheets(2).ListObjects(1)
    
    
    ArIn = InTbl.DataBodyRange
    ArStack = StackTbl.DataBodyRange
    
    For j = 1 To UBound(ArStack, 1)
    
        Stacks(j) = ArStack(j, 1)
        
    
    Next j
    
    For i = 1 To UBound(ArIn, 1)
    
        SplitIn = Split(ArIn(i, 1), " ")
        Qty = CLng(SplitIn(1))
        FromStack = CLng(SplitIn(3))
        ToStack = CLng(SplitIn(5))
        
        ModifyStacks Stacks, Qty, FromStack, ToStack
        
        
    
    Next i
    
    For Each Key In Stacks
    
        StackTops = StackTops & Mid(Stacks(Key), Len(Stacks(Key)), 1)
        
    
    Next Key
    
    MsgBox "Tops of all stacks are " & StackTops, vbOKOnly, "Stack Tops"
    
    
End Sub
Sub ModifyStacks(ByRef Stacks As Scripting.Dictionary, Qty As Long, FromStack As Long, ToStack As Long)

    Dim FirstStack As String
    Dim SecStack As String
    Dim MovedStack As String
    
    
    FirstStack = Stacks(FromStack)
    
    MovedStack = Mid(FirstStack, Len(FirstStack) - Qty + 1, Qty)
    
    SecStack = Stacks(ToStack) & MovedStack
    
    FirstStack = Mid(FirstStack, 1, Len(FirstStack) - Qty)
    
    
    Stacks(FromStack) = FirstStack
    Stacks(ToStack) = SecStack
    
        

End Sub
