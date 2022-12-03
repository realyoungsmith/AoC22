Sub StratGuideTotal()
    
    'for the first part of day 2, elves are playing rock paper scissors
    'they have a strat guide for some reason
    'first character is "the opponents" choice, second is mine
    'I will accumulate a score from each round
    'my choice + the outcome of the round = my score for the round
    
    Dim ChoiceDict As New Scripting.Dictionary
    Dim DataTbl As ListObject
    
    Dim ArData As Variant
    Dim Round As Variant
    
    Dim OppC As String
    Dim MyC As String
    
    Dim i As Long
    Dim RndScore As Long
    Dim TotalScore As Long
    
    
    
    ChoiceDict.Add Key:="A", Item:=1 'rock
    ChoiceDict.Add Key:="B", Item:=2 'paper
    ChoiceDict.Add Key:="C", Item:=3 'scissor
    ChoiceDict.Add Key:="X", Item:=1 'rock
    ChoiceDict.Add Key:="Y", Item:=2 'paper
    ChoiceDict.Add Key:="Z", Item:=3 'scissor
    
    
    Set DataTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArData = DataTbl.DataBodyRange
    RndScore = 0
    TotalScore = 0
    
    For i = 1 To UBound(ArData)
        
        Round = Split(ArData(i, 1))
        
        OppC = Round(0)
        MyC = Round(1)
            
        'same values always mean a tie
        If ChoiceDict(OppC) = ChoiceDict(MyC) Then
        
            RndScore = ChoiceDict(MyC) + 3 '3 points for a tie
        
        Else
        
            'lose cases
            If (ChoiceDict(OppC) = 1 And ChoiceDict(MyC) = 3) Or (ChoiceDict(OppC) = 2 And ChoiceDict(MyC) = 1) Or (ChoiceDict(OppC) = 3 And ChoiceDict(MyC) = 2) Then
                
                RndScore = ChoiceDict(MyC) + 0
                
            Else
            
                'should all be wins to get here
                RndScore = ChoiceDict(MyC) + 6
            End If
            
        End If
        
        TotalScore = TotalScore + RndScore
        
    Next i
    
    
    
    MsgBox "My TotalScore would be " & TotalScore, vbOKOnly, "Total Score"
    
    
    

End Sub

Sub StratGuideTotalFromOutCome()
    
    'for the first part of day 2, elves are playing rock paper scissors
    'they have a strat guide for some reason
    'first character is "the opponents" choice, second is mine
    'I will accumulate a score from each round
    'my choice + the outcome of the round = my score for the round
    
    Dim ChoiceDict As New Scripting.Dictionary
    Dim DataTbl As ListObject
    
    Dim ArData As Variant
    Dim Round As Variant
    
    Dim OppC As String
    Dim Outcome As String
    
    Dim i As Long
    Dim RndScore As Long
    Dim TotalScore As Long
    
    
    'in part two xyz represent the expected outcome
    ChoiceDict.Add Key:="A", Item:=1 'rock
    ChoiceDict.Add Key:="B", Item:=2 'paper
    ChoiceDict.Add Key:="C", Item:=3 'scissor
    ChoiceDict.Add Key:="X", Item:=0 'Lose
    ChoiceDict.Add Key:="Y", Item:=3 'Draw
    ChoiceDict.Add Key:="Z", Item:=6 'Win
    
    
    Set DataTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    ArData = DataTbl.DataBodyRange
    RndScore = 0
    TotalScore = 0
    
    For i = 1 To UBound(ArData)
        
        Round = Split(ArData(i, 1))
        
        OppC = Round(0)
        Outcome = Round(1)
            
        'same values always mean a tie
        If ChoiceDict(Outcome) = 3 Then
        
            RndScore = ChoiceDict(Outcome) + ChoiceDict(OppC) '3 points for a tie + whatever the opponents choice was
        
        Else
        
            'lose cases
            If ChoiceDict(Outcome) = 0 Then
            
                If ChoiceDict(OppC) = 1 Then
                    
                    RndScore = ChoiceDict(Outcome) + 3
                   
                Else
                    If ChoiceDict(OppC) = 2 Then
                    
                        RndScore = ChoiceDict(Outcome) + 1
                        
                    Else
                    
                        RndScore = ChoiceDict(Outcome) + 2
                    
                    End If
                    
                    
                End If
            Else
            
                'win cases
                If ChoiceDict(OppC) = 1 Then
                    
                    RndScore = ChoiceDict(Outcome) + 2
                Else
                
                    If ChoiceDict(OppC) = 2 Then
                    
                        RndScore = ChoiceDict(Outcome) + 3
                        
                    Else
                    
                        RndScore = ChoiceDict(Outcome) + 1
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
        TotalScore = TotalScore + RndScore
        
    Next i
    
    'should all be wins to get here
    'RndScore = ChoiceDict(Outcome) + 6
    
    MsgBox "My TotalScore would be " & TotalScore, vbOKOnly, "Total Score"
    
    
    

End Sub
