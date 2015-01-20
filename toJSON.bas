' toJSON  to convert the spreadsheet content into JSON-file with root node named "root"
Function toJSON(ByVal shtNm As String, ByVal keyEndCol As Integer, Optional ByVal fileName As String) As String
    ' KeyEndCol the first column without cellmerge

    
    Dim currentSht As Worksheet
    Set currentSht = ActiveSheet
    
    If shtNm <> "" Then
        Worksheets(shtNm).Activate
    End If
    
    Dim totalCol As Integer
    totalCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' load the key names
    Dim keyNm As Object
    Set keyNm = CreateObject("scripting.dictionary")
    
    Dim i As Integer
    For i = 1 To totalCol
        keyNm(i) = Array(Cells(1, i).Value, CStr(IsNumeric(Cells(2, i).Value)))
    Next i
    
    Dim j As Integer  ' row counter
    j = 2
    
    Dim y As Integer
    y = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim tmpRng As Range
    
    Dim res As String
    Dim res1 As String
    
    res1 = "{""root"":["
    
    i = 1
    Do While keyEndCol - i > 0
    
        
        Do While j <= y
            Set tmpRng = Cells(j, keyEndCol - i).MergeArea
            res = "{""" & keyNm(keyEndCol - i)(0) & """" & ":" & """" & Cells(j, keyEndCol - i).Value & """" & ",""Value""" & ":["
            ' for the rows in the merge
            For k = j To j + tmpRng.Rows.Count - 1
                res = res & packAttr(k, keyEndCol, totalCol, keyNm)
            Next k
            
            res = Left(res, Len(res) - 1) & "]},"
            
            res1 = res1 & res
            j = j + tmpRng.Rows.Count
        Loop
        

         i = i + 1
    Loop
    
    res1 = Left(res1, Len(res1) - 1) & "]}"
    
    
    
    currentSht.Activate
    
    If Not IsMissing(fileName) Then
        Dim fso As Object
        Set fso = CreateObject("scripting.filesystemobject")
        fso.createtextfile (ThisWorkbook.Path & "\" & fileName)
        Set f = fso.OpenTextFile(ThisWorkbook.Path & "\" & fileName, 2, True)
        
        f.write res1
        
    End If
    
    
    
    
    toJSON = res1
    
   
End Function

Private Function packAttr(ByVal targRow As Integer, ByVal startCol As Integer, ByVal endCol As Integer, ByVal dict As Object) As String

    Dim i As Integer
    
    Dim res As String
    res = "{"

    Dim arr
   
    
    For i = startCol To endCol
         arr = dict(i)
        
        If CBool(arr(1)) Then
            res = res & """" & arr(0) & """" & ":" & Application.WorksheetFunction.Substitute(Cells(targRow, i).Value, ",", ".") & ","
        Else
            res = res & """" & arr(0) & """" & ":" & """" & Trim(Cells(targRow, i).Value) & """" & ","
        End If

    Next i
    
    packAttr = Left(res, Len(res) - 1) & "},"

End Function


Private Function isMerged(ByVal rng As Range) As Boolean
    
    isMerged = rng.MergeCells

End Function


Private Sub toUnmerge()
    
    Dim cache As Range
    Dim c As Range

    For Each c In ActiveSheet.UsedRange.Cells
       If Not contains(cache, c) And isMerged(c) Then

       
        tmpValue = c.Value
        Set tmp = c.MergeArea
        tmp.UnMerge
        tmp.Value = tmpValue
        
        If cache Is Nothing Then
            Set cache = tmp
        Else
            Set cache = Application.Union(cache, tmp)
        End If
        
       End If
    Next c

End Sub


Private Sub toMerge()

    Dim c As Range
    
    Dim i As Integer
    Dim j As Integer


    i = 1
    j = 1
    
    For Each c In ActiveSheet.UsedRange.Cells
       If Not isMerged(c) And Not IsEmpty(c) Then

       
        tmpValue = c.Value
        
        ' last row of offset
        Do While c.Offset(i, 0).Value = tmpValue
            i = i + 1
        Loop
        
        i = i - 1
        
        ' last col of offset
        Do While c.Offset(0, j).Value = tmpValue
            j = j + 1
        Loop
        
        j = j - 1
        
        If i <> 0 Or j <> 0 Then
           With Range(c, c.Offset(i, j))
                .Value = ""
                .Merge
            End With
        End If
        
        c.Value = tmpValue
        
        i = 1
        j = 1
    
       End If
    Next c
    
End Sub


Private Function contains(rng1 As Range, ByVal rng2 As Range) As Boolean
    
    If rng1 Is Nothing Then
        contains = False
        Exit Function
    End If
    
    contains = Not (Application.Intersect(rng1, rng2) Is Nothing)
    
End Function
