Attribute VB_Name = "Module2"
Sub percent_change()
    For Each ws In Worksheets
        Dim worksheetName As String
        worksheetName = ws.Name
        
        Dim i As Long
        Dim j As Long
        Dim LastRowV As Long
        LastRowV = ws.Cells(Rows.Count, 11).End(xlUp).Row
        

        Dim percentchange As Double

        'Range("K2:K & LastRowV").NumberFormat = "General"
        
        i = 1
        j = 2
        
        For i = 2 To LastRowV
            If Cells(j, 3) <> 0 Then
            
            Cells(i, 11).Value = ((Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value)
            
            End If
        
          Next i
          
            
        
        
    Next
    
End Sub
