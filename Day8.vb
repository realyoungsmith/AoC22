Sub Day8()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    Dim CntVisibleOut As Long
    Dim X As Long
    Dim Y As Long
    Dim MaxX As Long
    Dim MaxY
    Dim Forrest As New Scripting.Dictionary
    Dim Trees As Variant
    
    
    
    
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    
    ArIn = TblIn.ListColumns(1).DataBodyRange
    
    CntVisibleOut = 0
    
    'This is left to right
    'row by row
    '
    MaxX = 0
    MaxY = 0
    For Y = 1 To UBound(ArIn)
        For X = 1 To Len(ArIn(Y, 1))
            
            If Not Forrest.Exists(X & "," & Y) Then
             
                'north/south edge
                If Y = 1 Or Y = UBound(ArIn) Then
                
                    Forrest(X & "," & Y) = "Visible"
                    If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxX Then

                        MaxX = CLng(Mid(ArIn(Y, 1), X, 1))
                    
                    End If
                Else
                    'west/east edge
                    If X = 1 Or X = Len(ArIn(Y, 1)) Then
                      
                      
                        Forrest(X & "," & Y) = "Visible"
                        If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxX Then

                            MaxX = CLng(Mid(ArIn(Y, 1), X, 1))
                    
                        End If
                    Else
                    
                        'Everything else
                        '
                        If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxX Then
                        
                            Forrest(X & "," & Y) = "Visible"
                            
                             
'                            If ArIn(Y, 1) > MaxY Then
'                                MaxY = ArIn(Y, 1)
'                            End If
                            

                                MaxX = CLng(Mid(ArIn(Y, 1), X, 1))
                           
                        End If
                        
                    End If
                    
                        
                    
                End If
            End If
        Next X
        MaxX = 0
        
    Next Y
    
    Trees = Split(Join(Forrest.Items(), "|"), "|")
    MaxX = 0
    MaxY = 0
    
    For Y = 1 To UBound(ArIn)
        For X = Len(ArIn(Y, 1)) To 1 Step -1
            
            If Not Forrest.Exists(X & "," & Y) Then
             
                If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxX Then
                
                    Forrest(X & "," & Y) = "Visible"

                    MaxX = CLng(Mid(ArIn(Y, 1), X, 1))
                   
                End If
                
                
                    
            Else
                
                If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxX Then

                        MaxX = CLng(Mid(ArIn(Y, 1), X, 1))
                    
                End If
                
            End If
        Next X
        MaxX = 0
    Next Y
    Trees = Split(Join(Forrest.Items(), "|"), "|")
    
    MaxX = 0
    MaxY = 0
    
    For X = 1 To Len(ArIn(1, 1))
        For Y = 1 To UBound(ArIn)
            If Not Forrest.Exists(X & "," & Y) Then
               
                If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxY Then
                
                    Forrest(X & "," & Y) = "Visible"
                    
                    MaxY = CLng(Mid(ArIn(Y, 1), X, 1))
                   
                End If
                  
                
            Else
                If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxY Then
                
                    MaxY = CLng(Mid(ArIn(Y, 1), X, 1))
                
              
                End If
            End If
        Next Y
        MaxY = 0
    Next X
      
    MaxX = 0
    MaxY = 0
    Trees = Split(Join(Forrest.Items(), "|"), "|")
    
    X = 1
    
    For X = 1 To Len(ArIn(1, 1))
        For Y = UBound(ArIn) To 1 Step -1
        
            If Not Forrest.Exists(X & "," & Y) Then
                
                If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxY Then
                
                    Forrest(X & "," & Y) = "Visible"
                    
                    MaxY = CLng(Mid(ArIn(Y, 1), X, 1))
                   
                End If
                
            Else
                If CLng(Mid(ArIn(Y, 1), X, 1)) > MaxY Then
                
                    MaxY = CLng(Mid(ArIn(Y, 1), X, 1))
                
              
                End If
            End If
        Next Y
        MaxY = 0
    Next X
    Trees = Split(Join(Forrest.Items(), "|"), "|")
    
    MsgBox "Tree Count is " & UBound(Trees) + 1, vbOKOnly, "Tree count"
    
End Sub
Sub Day8pt2()
    
    Dim TblIn As ListObject
    Dim ArIn As Variant
    Dim CntVisibleOut As Long
    Dim X As Long
    Dim Y As Long
    Dim MaxXP, MaxYP, MaxXN, MaxYN As Long
    Dim XP, YP, XN, YN As Long
    Dim Forrest As New Scripting.Dictionary
    Dim Trees As Variant
    Dim CTree As Long
    Dim TTree As Long
    
    
    
    
    
    Set TblIn = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    
    ArIn = TblIn.ListColumns(1).DataBodyRange
    
    CntVisibleOut = 0
    
    'This is left to right
    'row by row
    '
    MaxXP = 0
    MaxYP = 0
    MaxXN = 0
    MaxYN = 0
    
    For Y = 1 To UBound(ArIn)
        For X = 1 To Len(ArIn(Y, 1))
            XP = 1
            YP = 1
            XN = 1
            YN = 1
            
            If Y = 16 And X = 50 Then
                Y = Y
            End If
            CTree = CLng(Mid(ArIn(Y, 1), X, 1))
            If X > 1 And Y > 1 And X < Len(ArIn(Y, 1)) And Y < UBound(ArIn) Then
            
                While CTree > CLng(Mid(ArIn(Y, 1), X + XP, 1)) And X + XP < Len(ArIn(Y, 1))
                
                
                
                    XP = XP + 1
                Wend
                
                
                While CTree > CLng(Mid(ArIn(Y, 1), X - XN, 1)) And X - XN > 1
                
                
                
                    XN = XN + 1
                Wend
                
                While CTree > CLng(Mid(ArIn(Y + YP, 1), X, 1)) And Y + YP < UBound(ArIn)
                
                
                
                    YP = YP + 1
                Wend
                
                While CTree > CLng(Mid(ArIn(Y - YN, 1), X, 1)) And Y - YN > LBound(ArIn)
                
                
                
                    YN = YN + 1
                Wend
                MaxY = XP * XN * YP * YN
                Forrest(X & "," & Y) = MaxY
                
            End If
            
        Next X
  
    Next Y
    
    
    Trees = Split(Join(Forrest.Items(), "|"), "|")
    
    MaxY = 0
    For j = 0 To UBound(Trees)
        
        If CLng(Trees(j)) > MaxY Then
            MaxY = Trees(j)
        End If
        
    
    Next j
    
    MsgBox "Max View " & MaxY, vbOKOnly, "Max View"
    
    
End Sub

