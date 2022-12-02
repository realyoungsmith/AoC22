'file says .vb just so git would be happy but this was all VBA Excel
Sub elfwiththemost()
    
    'lots of elves carrying snacks
    'each group of numbers is an elf, each number is a snack, the value of the snack is its calories
    'use data set see the highest amount of calories carried by a single elf
    
    Dim TblInput As Excel.ListObject
    Dim ArInput As Variant
    
    Dim CCal As Double
    Dim MCal As Double
    
    Dim i As Long
    
    Set TblInput = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArInput = TblInput.DataBodyRange
    CCal = 0
    MCal = 0
    i = 1
    
    While i < UBound(ArInput)
        
        If ArInput(i, 1) <> "" Then
        
            CCal = CCal + ArInput(i, 1)
        Else
            'if the current elf has more callories than what we've seen, set it as the new max
            If CCal > MCal Then
        
                MCal = CCal
                
            End If
            CCal = 0
        End If
        
        
        
        i = i + 1
    
    Wend
    
    MsgBox "Max Calorie is " & MCal, vbOKOnly, "Maximum Calories"

End Sub
Sub Top3withthemost()
    
    'lots of elves carrying snacks
    'each group of numbers is an elf, each number is a snack, the value of the snack is its calories
    'use data set see the highest amount of calories carried by a single elf
    'for this second part, find the top three elves
    
    
    Dim TblInput As Excel.ListObject
    Dim TblOutput As Excel.ListObject
    
    Dim Rng As Range
    
    Dim ArInput As Variant
    Dim ArOutput As Variant
    
    Dim Swapped As Boolean
    Dim CCal As Double
    Dim MCal As Double
    
    Dim i As Long
    
    Set TblInput = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    
    ArInput = TblInput.DataBodyRange
    ReDim ARThree(1 To 4) 'oversized array to keep top 3 in
    
    CCal = 0
    MCal = 0
    i = 1
    
    While i < UBound(ArInput)
        
        If ArInput(i, 1) <> "" Then
        
            CCal = CCal + ArInput(i, 1)
        Else
            For j = 1 To 3
    
                If CCal > ARThree(j) Then
                    
                    'value greater than current line, starting from
                    'once we find a top 3 value less than our current, shift it all down and insert the current
                    For k = 4 To j Step -1
                        
                        ARThree(k) = ARThree(j)
                        
                    Next k

                    ARThree(j) = CCal
                      
                    Exit For
                  
                End If
 
            Next j
            
            CCal = 0
        End If

        i = i + 1
    
    Wend
    
    'sum up the top three
    For j = 1 To 3
    
        MCal = MCal + CDbl(ARThree(j))
        
    Next j

    MsgBox "Max Calorie is " & MCal, vbOKOnly, "Maximum Calories"

End Sub
