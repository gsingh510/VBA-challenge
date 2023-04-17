Attribute VB_Name = "Module3"
Sub totalstock()

    For Each ws In Worksheets
        Dim worksheetName As String
        worksheetName = ws.Name

        Dim i As Long
        Dim j As Long
        
        Dim LastRowO As Long
        Dim LastRowA As Long
        Dim GI As Double
        Dim GD As Double
        Dim GV As Double
        
        counter = 2
        j = 2
        LastRowO = ws.Cells(Rows.Count, 12).End(xlUp).Row
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowA
                ws.Cells(counter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                counter = counter + 1
                
                j = i + 1
            Next i
             
             
GI = ws.Range("Q2").Value
GD = ws.Range("Q3").Value
GV = ws.Range("Q4").Value
    
    For i = 2 To LastRowO
    

    
        
        
    Next ws

End Sub
