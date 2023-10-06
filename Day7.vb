Sub day7()

    
    Dim PathDict As New Scripting.Dictionary
    Dim DirDict As New Scripting.Dictionary
    
    Dim InTbl As ListObject
    
    Dim ArIn As Variant
    Dim SplitFile As Variant
    Dim ArSizes As Variant
    Dim SplitStr As Variant
    
    Dim SplitPath As String
    Dim CPath As String
    Dim CDir As String
    
    Dim CDirSize As Long
    Dim TotalMem As Long
    Dim SplitPos As Long
    Dim PathLength As Long
    
    Dim Total As Long
    Dim Free As Long
    Dim Used As Long
    Dim Diff As Long
    
    
    
    Set InTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    
    ArIn = InTbl.DataBodyRange
    
    
    
    For i = 1 To UBound(ArIn)
        
       
        If InStr(ArIn(i, 1), "$ ") Then
            
            ArIn(i, 1) = Replace(ArIn(i, 1), "$ ", "")
        End If
        
            
        SplitStr = Split(ArIn(i, 1))
        
        Select Case SplitStr(0)
        
        Case "cd"
            
            If SplitStr(1) = "/" Then
            
               'start over
               CPath = "/"
            
            Else
            
                If SplitStr(1) = ".." Then
                    
                    'go back one
                    SplitPos = InStrRev(CPath, "/")
                    
                    CPath = Left(CPath, SplitPos - 1)
                    
                Else
                
                    
                    'go in lvl
                    If CPath <> "/" Then
                    
                        CPath = CPath & "/" & SplitStr(1)
                    Else
                    
                        CPath = CPath & SplitStr(1)
                    End If
                    
                    If Not PathDict.Exists(CPath) Then
                    
                    
                        PathDict(CPath) = 0
                    End If
                    
                
                End If
                
                
                
            End If
        Case "dir"
            
            'do nothing also?
            
        
        Case "ls"
            
            'nothing technically
            
            
        Case Else
            
            'should be files at this point
            'account for parent directory(ies)
            
            PathDict(CPath) = PathDict(CPath) + CLng(SplitStr(0))
            SplitPath = CPath
            
           
            
                
                PathLength = (Len(CPath) - Len(Replace(CPath, "/", "")))    'this is just the length of the path minus length of the path without "/", number of times to iterate
                '/something/something/dsoemfhd/sfjodsf/fdjskla
                While PathLength > 1    'Plus one i think since all paths start with /
                
                    SplitPos = InStrRev(SplitPath, "/")
                        
                    SplitPath = Left(SplitPath, SplitPos - 1)
                    
                    PathDict(SplitPath) = PathDict(SplitPath) + CLng(SplitStr(0))
                    PathLength = PathLength - 1
                    
                Wend
                
         
            
        
        End Select
        
        
        
    
    Next i
    
   
    
    Arsize = Split(Join(PathDict.Items, "|"), "|")
    ArDir = Split(Join(PathDict.Keys, "|"), "|")
    
    For i = LBound(Arsize) To UBound(Arsize)
        
        If CLng(Arsize(i)) <= 100000 Then
            
            TotalMem = TotalMem + CLng(Arsize(i))
            
            
        End If
        
              
    Next i
    
    MsgBox "Total Mem " & TotalMem, vbOKOnly, "Total Mem"
    
End Sub
Sub day7pt2()

    
    Dim PathDict As New Scripting.Dictionary
    Dim DirDict As New Scripting.Dictionary
    
    Dim InTbl As ListObject
    
    Dim ArIn As Variant
    Dim SplitFile As Variant
    Dim ArSizes As Variant
    Dim SplitStr As Variant
    
    Dim SplitPath As String
    Dim CPath As String
    Dim CDir As String
    
    Dim CDirSize As Long
    Dim TotalMem As Long
    Dim SplitPos As Long
    Dim PathLength As Long
    
    Dim Total As Long
    Dim Free As Long
    Dim Used As Long
    Dim Diff As Long
    Dim Min As Long
    
    
    
    
    Set InTbl = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    
    ArIn = InTbl.DataBodyRange
    
    
    
    For i = 1 To UBound(ArIn)
        
       
        If InStr(ArIn(i, 1), "$ ") Then
            
            ArIn(i, 1) = Replace(ArIn(i, 1), "$ ", "")
        End If
        
            
        SplitStr = Split(ArIn(i, 1))
        
        Select Case SplitStr(0)
        
        Case "cd"
            
            If SplitStr(1) = "/" Then
            
               'start over
               CPath = "/"
            
            Else
            
                If SplitStr(1) = ".." Then
                    
                    'go back one
                    SplitPos = InStrRev(CPath, "/")
                    
                    CPath = Left(CPath, SplitPos - 1)
                    
                Else
                
                    
                    'go in lvl
                    If CPath <> "/" Then
                    
                        CPath = CPath & "/" & SplitStr(1)
                    Else
                    
                        CPath = CPath & SplitStr(1)
                    End If
                    
                    If Not PathDict.Exists(CPath) Then
                    
                    
                        PathDict(CPath) = 0
                    End If
                    
                
                End If
                
                
                
            End If
        Case "dir"
            
            'do nothing also?
            
        
        Case "ls"
            
            'nothing technically
            
            
        Case Else
            
            'should be files at this point
            'account for parent directory(ies)
            
            PathDict(CPath) = PathDict(CPath) + CLng(SplitStr(0))
            SplitPath = CPath
            
           
            
                
                PathLength = (Len(CPath) - Len(Replace(CPath, "/", "")))    'this is just the length of the path minus length of the path without "/", number of times to iterate
                '/something/something/dsoemfhd/sfjodsf/fdjskla
                While PathLength > 0 And SplitPath <> "/"    'Plus one i think since all paths start with /
                
                    SplitPos = InStrRev(SplitPath, "/")
                        
                    If SplitPos = 1 Then
                        SplitPath = "/"
                    Else
                        SplitPath = Left(SplitPath, SplitPos - 1)
                    End If
                    
                    PathDict(SplitPath) = PathDict(SplitPath) + CLng(SplitStr(0))
                    PathLength = PathLength - 1
                    
                Wend
                
         
            
        
        End Select
        
        
        
    
    Next i
    
    Total = 70000000
    Used = CLng(PathDict("/"))
    
    Free = Total - Used
    Diff = 30000000
    Neeeded = 30000000
    Min = Total
    For Each Key In PathDict
    
        If PathDict(Key) + Free >= Neeeded And PathDict(Key) < Min Then
            
            Min = PathDict(Key)
        End If
        
    
    Next Key
    
    
'    Arsize = Split(Join(PathDict.Items, "|"), "|")
'
'    For i = LBound(Arsize) To UBound(Arsize)
'
'        If CLng(Arsize(i)) <= 100000 Then
'
'            TotalMem = TotalMem + CLng(Arsize(i))
'
'
'        End If
'
'
'    Next i
    
    MsgBox "Total Mem " & Min, vbOKOnly, "Total Mem"
    
End Sub
