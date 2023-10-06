Sub Part1()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    
    Dim Monkeys As New Scripting.Dictionary
    Dim MonkeyBusiness As New Scripting.Dictionary
    Dim NewMonk As New Scripting.Dictionary
    Dim Monkey As New Scripting.Dictionary
    Dim Test As New Scripting.Dictionary
    
    Dim MonkeyNumber As String
    Dim NewMonkey As String
    Dim ArItems As Variant
    Dim ArOp As Variant
    Dim ArHold As Variant
    
    Dim StrTest As String
    Dim StrIn As String
    
    Dim old As Long
    Dim NewNew As Long
    Dim Worry As Long
    Dim Round As Long
    Dim Monkone As Long
    Dim Monktwo As Long
    
    
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = TblIn.DataBodyRange
    
    ReDim ARTest(1 To 2)
    
    'build monkeys
    'each monkey contains 3 dictionairies
    'first is Items, array of current worry for it's items
    'second is that monkeys operation as an array
    'third is the test, key will be the test, array of items is the result, true being the first in the array
    
    
    For i = 1 To UBound(ArIn, 1)
        
        StrIn = ArIn(i, 1)
        
        If StrIn <> "" Then
              
            If InStr(StrIn, "Monkey") > 0 Then
                MonkeyNumber = GetMonkeyNumber(StrIn)
            End If
            
            If InStr(StrIn, "items:") > 0 Then
                ArItems = GetItems(StrIn)
            End If
           
            If InStr(StrIn, "Operation") > 0 Then
                ArOperation = GetOperation(StrIn)
            End If
           
            If InStr(StrIn, "Test") > 0 Then
                StrTest = GetTest(StrIn)
            End If
            
            If InStr(StrIn, "true") > 0 Then
                ARTest(1) = GetTrue(StrIn)
            End If
            
            If InStr(StrIn, "false") > 0 Then
                ARTest(2) = GetFalse(StrIn)
            End If
            
              
        Else
            'when hitting a blank, use all the info gathered to build the monkey
            
            Test(StrTest) = ARTest
            
            Monkey("Items") = ArItems
            Monkey("Operation") = ArOperation
            Set Monkey("Test") = Test
            
            Set Monkeys(CStr(MonkeyNumber)) = Monkey
            
            
            Set Monkey = New Dictionary
            Set Items = New Dictionary
            Set Operation = New Dictionary
            Set Test = New Dictionary
        End If
       
    Next i
    
    'monkeys built time to start monkeying around
    Round = 0
    While Round < 20
    
        For Each Key In Monkeys
            
            Set Monkey = Monkeys(Key)
            ArItems = Monkey("Items")
            ArOperation = Monkey("Operation")
            Set Test = Monkey("Test")
            
            'inspect items
            If Not IsEmpty(ArItems) Then
            
                For i = 0 To UBound(ArItems)
                
                    'do operation on item
                    
                    old = ArItems(i)
                    
                    NewNew = DoOp(ArOperation, old)
                    If Not MonkeyBusiness.Exists(Key) Then
                    
                        MonkeyBusiness(Key) = 1
                    Else
                        
                        MonkeyBusiness(Key) = MonkeyBusiness(Key) + 1
                    
                    End If
                    
                    'i'm so glad the aids monkey didn't break my christmas gear
                    NewNew = CLng(Application.WorksheetFunction.RoundDown(CDbl(NewNew) / 3, 0))
                    
                    ARTest = Test(Test.Keys(0))
                    
                    'test new item
                    If NewNew Mod CLng(Test.Keys(0)) = 0 Then
                        
                        NewMonkey = ARTest(1)
                        
                    
                    Else
                    
                        NewMonkey = ARTest(2)
                    
                    End If
                    If NewNew = 1580 Then
                        Key = Key
                    End If
                    
                    
                    If NewMonkey = "0" Then
                   
                        Key = Key
                    
                    End If
                    
                    
                    Set NewMonk = Monkeys(CStr(NewMonkey))
                    
                    ArHold = NewMonk("Items")
                    
                    If IsEmpty(ArHold) Then
                        
                        ReDim ArHold(0 To 0)
                        ArHold(UBound(ArHold)) = NewNew
                    Else
                        ReDim Preserve ArHold(0 To UBound(ArHold) + 1)
                    
                        ArHold(UBound(ArHold)) = NewNew
                    End If

                    NewMonk("Items") = ArHold
                    
                    Set Monkeys(NewMonkey) = NewMonk
                    
                    Set NewMonk = New Dictionary
                               
                    
                Next i
                
                'after throwing all items clear monkey items ar
                'how the fuck can i clear a god damn array fucking stupid as monkey fucing bullshit ass fucking  microsoft, why the fuck can arrays only be = Empty when you first declare the god damn
                'mother fucking things, what kind of stupid as fucking bullshit is this, i'm going to need another sub to create ANOTHER array that is empty just to put in this mother fucking
                'stupid ass fucking monkey bullshit
                EmptyGodDamnArray Monkey
                
                
                
               
                
                Set Monkeys(Key) = Monkey
                
                Set Monkey = New Dictionary

            End If
            
        Next Key
        Round = Round + 1
        
        For Each Key In Monkeys
            Set Monkey = Monkeys(Key)
            ArItems = Monkey("Items")
            Set Monkey = New Dictionary
            
        Next Key
        
    Wend
    
    Monkone = 0
    Monktwo = 0
    
    For Each Key In MonkeyBusiness
        
        If MonkeyBusiness(Key) > Monkone Then
        
            If Monkone > Monktwo Then
                
                Monktwo = Monkone
                Monkone = MonkeyBusiness(Key)
            Else
                Monkone = MonkeyBusiness(Key)
            End If
            
        End If
        
    Next Key
    
    MsgBox Monkone * Monktwo, vbOKOnly, "Level of Simian shenanigans"
    

End Sub
Sub Part2()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    
    Dim Dict As Scripting.Dictionary
    
    
    Set TblIn = ThisWorkbook
    
    'I could never figure a way to handle the extremley large numbers generated in the second part of this challenge
    'It was about 4am and I was just a BIT frustrated if you couldn't tell haha
    

End Sub

Function GetMonkeyNumber(Line As String) As String

    Dim MonkeySplit As Variant
    
    MonkeySplit = Split(Replace(Line, ":", ""), " ")
    
    GetMonkeyNumber = MonkeySplit(1)
    

End Function
Function GetItems(Line As String) As Variant

    Dim LineSplit As Variant
    Dim ItemSplit As Variant
    
    LineSplit = Split(Line, ": ")
    ItemSplit = Split(LineSplit(1), ", ")
    
    GetItems = ItemSplit
    
    
    
End Function
Function GetOperation(Line As String) As Variant

    Dim LineSplit As Variant
    Dim OpSplit As Variant
    
    LineSplit = Split(LTrim(Line), ": ")
    OpSplit = Split(LineSplit(1), " ")
    
    GetOperation = OpSplit
    
End Function
Function GetTest(Line As String) As String

    Dim LineSplit As Variant
    'monkeys always divide so
    Dim TestSplit As Variant
    
    
    LineSplit = Split(Line, ": ")
    TestSplit = Split(LineSplit(1), " ")
    
    GetTest = TestSplit(UBound(TestSplit))
    
    
End Function
Function GetTrue(Line As String) As String

    Dim LineSplit As Variant
    
    LineSplit = Split(LTrim(Line), " ")
    
    GetTrue = LineSplit(UBound(LineSplit))
    
End Function
Function GetFalse(Line As String) As String

    Dim LineSplit As Variant
    
    LineSplit = Split(LTrim(Line), " ")
    
    GetFalse = LineSplit(UBound(LineSplit))
    
End Function
Function DoOp(ArOperation As Variant, old As Long) As Long
    Dim ArOpHold
    
    ArOpHold = ArOperation
    For i = 1 To UBound(ArOpHold)
        
        If ArOpHold(i) = "old" Then
            
            ArOpHold(i) = old
        End If
    Next i
    
    If ArOperation(3) = "+" Then
        DoOp = old + CLng(ArOpHold(4))
    Else
        If ArOperation(3) = "*" Then
            DoOp = old * CLng(ArOpHold(4))
        End If
    End If
    
    'DoOp = Application.Evaluate(old & ArOperation(3) & ArOperation(4))
    

End Function
Sub EmptyGodDamnArray(ByRef Monkey As Variant)
    
    
    Dim AREmpty As Variant
    
    Monkey("Items") = AREmpty
    
    
    
    
    
End Sub
