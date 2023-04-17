Attribute VB_Name = "Module1"
Sub Symbol()
    For Each ws In Worksheets
        Dim worksheetName As String
        worksheetName = ws.Name
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest Percent Increase"
        ws.Range("O3").Value = "Greatest Percent Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Columns("J").ColumnWidth = 18
        ws.Columns("K").ColumnWidth = 18
        ws.Columns("L").ColumnWidth = 18
        ws.Columns("O").ColumnWidth = 20
      Next
End Sub


Sub Ticker_Year_Volume()
    For Each ws In Worksheets
        Dim wsName As String
        wsName = ws.Name
        Dim i As Long
        Dim j As Long
        Dim Tick As Long
        Tick = 1
        j = 2
        Dim LastX As Long
        LastX = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To LastX
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Tick = Tick + 1
                ws.Cells(Tick, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(Tick, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
                    If ws.Cells(Tick, 10).Value < 0 Then
                    
                    ws.Cells(Tick, 10).Interior.Color = vbRed
                    
                    Else
                    
                    ws.Cells(Tick, 10).Interior.Color = vbYellow
                    
                    End If
                                  
                End If
            Next i
    Next ws
End Sub
