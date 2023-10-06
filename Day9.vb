Sub Part1()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    
    Dim Visited As New Scripting.Dictionary
    Dim Directions As New Scripting.Dictionary
    
    Directions("U") = 1
    Directions("D") = -1
    Directions("R") = 1
    Directions("L") = -1
    
    Dim Dir As String
        
    Dim Dis As Long
    Dim i As Long
    
    Dim Rope As New Scripting.Dictionary
    Dim Location As Variant
    Dim PriorLocation As Variant
    Dim ArVisited As Variant
    
    ReDim Location(1 To 2)
    ReDim PriorLocation(1 To 2)
    
    
    '1 = head, 2 = tail
    For i = 1 To 2
        
        Location(1) = 0
        Location(2) = 0
        
        Rope(i) = Location
    
    Next i
    
    
    Dim SplitIn As Variant
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = TblIn.DataBodyRange
    
    For i = 1 To UBound(ArIn)
    
        SplitIn = Split(ArIn(i, 1), " ")
        Dir = SplitIn(0)
        Dis = CLng(SplitIn(1))
        For j = 1 To Dis
        
            For Each Key In Rope
                
                Select Case Key
                
                Case 1
                
                    Select Case Dir
                    
                    Case "L"
                        Location = Rope(Key)
                        
                        Location(1) = Location(1) - 1
                        
                        Rope(Key) = Location
                    Case "R"
                        Location = Rope(Key)
                        
                        Location(1) = Location(1) + 1
                        
                        Rope(Key) = Location
                    Case "D"
                        Location = Rope(Key)
                        
                        Location(2) = Location(2) - 1
                        
                        Rope(Key) = Location
                    Case "U"
                        Location = Rope(Key)
                        
                        Location(2) = Location(2) + 1
                        
                        Rope(Key) = Location
                
                    End Select
                Case Else
                    PriorLocation = Rope(CLng(Key) - 1)
                    Location = Rope(Key)
                    'x, y
                    
                    If (Abs((PriorLocation(1)) - Location(1)) = 1 And Abs(PriorLocation(2) - Location(2)) > 1) Or (Abs(PriorLocation(2) - Location(2)) = 1 And Abs((PriorLocation(1)) - Location(1)) > 1) Then
                        'new location = current location
                        Location(1) = Location(1) + (Abs((PriorLocation(1)) - Location(1)) / (PriorLocation(1) - Location(1)))
                        Location(2) = Location(2) + (Abs((PriorLocation(2)) - Location(2)) / (PriorLocation(2) - Location(2)))
                    Else
                        
                        If Abs(PriorLocation(1) - Location(1)) > 1 Then
                            Location(1) = Location(1) + (Abs((PriorLocation(1)) - Location(1)) / (PriorLocation(1) - Location(1)))
                        End If
                        
                        If Abs(PriorLocation(2) - Location(2)) > 1 Then
                            Location(2) = Location(2) + (Abs((PriorLocation(2)) - Location(2)) / (PriorLocation(2) - Location(2)))
                        End If
                        
                    End If
                    
                    Rope(Key) = Location
                    
                    
                    If Not Visited.Exists(Location(1) & "," & Location(2)) Then
                        Visited(Location(1) & "," & Location(2)) = "Visited"
                    End If
                    
                        
                End Select
                
            Next Key
       Next j
       
    Next i
    
    ArVisited = Split(Join(Visited.Items(), "|"), "|")
    
    MsgBox "Places " & UBound(ArVisited) + 1, vbOKOnly, "Places"
    

End Sub
Sub Part2()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    
    Dim Visited As New Scripting.Dictionary
    Dim Directions As New Scripting.Dictionary
    
    Directions("U") = 1
    Directions("D") = -1
    Directions("R") = 1
    Directions("L") = -1
    
    Dim Dir As String
        
    Dim Dis As Long
    Dim i As Long
    
    Dim Rope As New Scripting.Dictionary
    Dim Location As Variant
    Dim PriorLocation As Variant
    Dim ArVisited As Variant
    
    ReDim Location(1 To 2)
    ReDim PriorLocation(1 To 2)
    
    
    '1 = head, 2 = tail
    For i = 1 To 10
        
        Location(1) = 0
        Location(2) = 0
        
        Rope(i) = Location
    
    Next i
    
    
    Dim SplitIn As Variant
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = TblIn.DataBodyRange
    
    For i = 1 To UBound(ArIn)
    
        SplitIn = Split(ArIn(i, 1), " ")
        Dir = SplitIn(0)
        Dis = CLng(SplitIn(1))
        For j = 1 To Dis
        
            For Each Key In Rope
                
                Select Case Key
                
                Case 1
                
                    Select Case Dir
                    
                    Case "L"
                        Location = Rope(Key)
                        
                        Location(1) = Location(1) - 1
                        
                        Rope(Key) = Location
                    Case "R"
                        Location = Rope(Key)
                        
                        Location(1) = Location(1) + 1
                        
                        Rope(Key) = Location
                    Case "D"
                        Location = Rope(Key)
                        
                        Location(2) = Location(2) - 1
                        
                        Rope(Key) = Location
                    Case "U"
                        Location = Rope(Key)
                        
                        Location(2) = Location(2) + 1
                        
                        Rope(Key) = Location
                
                    End Select
                Case Else
                    PriorLocation = Rope(CLng(Key) - 1)
                    Location = Rope(Key)
                    'x, y
                   
                    If (Abs((PriorLocation(1)) - Location(1)) = 1 And Abs(PriorLocation(2) - Location(2)) > 1) Or (Abs(PriorLocation(2) - Location(2)) = 1 And Abs((PriorLocation(1)) - Location(1)) > 1) Then
                        'new location = current location
                        Location(1) = Location(1) + (Abs((PriorLocation(1)) - Location(1)) / (PriorLocation(1) - Location(1)))
                        Location(2) = Location(2) + (Abs((PriorLocation(2)) - Location(2)) / (PriorLocation(2) - Location(2)))
                    Else
                        
                        If Abs(PriorLocation(1) - Location(1)) > 1 Then
                            Location(1) = Location(1) + (Abs((PriorLocation(1)) - Location(1)) / (PriorLocation(1) - Location(1)))
                        End If
                        
                        If Abs(PriorLocation(2) - Location(2)) > 1 Then
                            Location(2) = Location(2) + (Abs((PriorLocation(2)) - Location(2)) / (PriorLocation(2) - Location(2)))
                        End If
                        
                    End If
                    
                    Rope(Key) = Location
                    
                    If Key = 10 Then
                    
                        If Not Visited.Exists(Location(1) & "," & Location(2)) Then
                            Visited(Location(1) & "," & Location(2)) = "Visited"
                        End If
                    End If
                        
                    
                    
                        
                End Select
                
            Next Key
       Next j
    ArVisited = Split(Join(Visited.Items(), "|"), "|")
    Next i
    
    ArVisited = Split(Join(Visited.Items(), "|"), "|")
    
    MsgBox "Places " & UBound(ArVisited) + 1, vbOKOnly, "Places"
    
    

End Sub

Sub TestCode()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    
    Dim Visited As New Scripting.Dictionary
    
        
    Dim Dir As String
        
    Dim Dis As Long
    Dim i As Long
    
    Dim Rope As New Scripting.Dictionary
    Dim Location As Variant
    Dim PriorLocation As Variant
    Dim ArVisited As Variant
    
    ReDim Location(1 To 2)
    ReDim PriorLocation(1 To 2)
    
    
    '1 = head, 2 = tail
    For i = 1 To 2
        
        Location(1) = 0
        Location(2) = 0
        
        Rope(i) = Location
    
    Next i
    
    
    Dim SplitIn As Variant
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArIn = TblIn.DataBodyRange
    
    For i = 1 To UBound(ArIn)
    
        SplitIn = Split(ArIn(i, 1), " ")
        Dir = SplitIn(0)
        Dis = CLng(SplitIn(1))
        
            For Each Key In Rope
                
                Select Case Key
                
                Case 1
                
                    Select Case Dir
                    
                    Case "L"
                        Location = Rope(Key)
                        
                        Location(1) = Location(1) - Dis
                        
                        Rope(Key) = Location
                    Case "R"
                        Location = Rope(Key)
                        
                        Location(1) = Location(1) + Dis
                        
                        Rope(Key) = Location
                    Case "D"
                        Location = Rope(Key)
                        
                        Location(2) = Location(2) - Dis
                        
                        Rope(Key) = Location
                    Case "U"
                        Location = Rope(Key)
                        
                        Location(2) = Location(2) + Dis
                        
                        Rope(Key) = Location
                
                    End Select
                Case Else
                    PriorLocation = Rope(CLng(Key) - 1)
                    Location = Rope(Key)
                    'x, y
                    
                    If Abs((PriorLocation(1)) - Location(1)) > 1 Then
                        'new location = current location
                        Location(1) = Location(1) + (PriorLocation(1) - Location(1)) - (Abs((PriorLocation(1)) - Location(1)) / PriorLocation(1) - Location(1))
                                   
                    End If
                    
                    If Abs(PriorLocation(2) - Location(2)) > 1 Then
                        Location(2) = Location(2) + (PriorLocation(2) - Location(2)) - (Abs((PriorLocation(2)) - Location(2)) / PriorLocation(2) - Location(2))
                    End If
                    
                    If Not Visited.Exists(Location(1) & "," & Location(2)) Then
                        Visited(Location(1) & "," & Location(2)) = "Visited"
                    End If
                    
                        
                End Select
                
            Next Key
       
    Next i
    
    ArVisited = Split(Join(Visited.Items(), "|"), "|")
    
    MsgBox "Places " & UBound(ArVisited) + 1, vbOKOnly, "Places"
    

End Sub
