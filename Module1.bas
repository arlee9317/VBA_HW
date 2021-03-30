Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim Ticker As String
Dim Percent_change As Double
Dim Yearly_change As Double
Dim Volume As Double
Dim x1 As Double
Dim x2 As Double
Dim summary_table_row As Integer
Dim ws As Worksheet


For Each ws In ActiveWorkbook.Worksheets
    'delcaring variables
    Volume = 0
    x1 = ws.Cells(2, 3).Value
    summary_table_row = 2
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'setting each column header to autofill and center
    ws.Range("I1:L1").EntireColumn.AutoFit
    ws.Range("I1:L1").HorizontalAlignment = xlCenter
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    
    
    For i = 2 To Last_Row
        
        'Comparing if ticker symbol equals the next row ticker symbol.
        'If it is then place ticker symbol somewhere on excel page.
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            'accounting for the last row in the ticker range to be added in
            Volume = Volume + Cells(i, 7).Value
            ws.Cells(summary_table_row, 12).Value = Volume
            
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(summary_table_row, 9).Value = Ticker
        
            'yearly change calc for each ticker value
            x2 = ws.Cells(i, 6).Value
            Yearly_change = x2 - x1
            ws.Cells(summary_table_row, 10).Value = Yearly_change
        
            'Percent change calc for each ticker value, this also makes sure to account for divide by zeros
            If x1 <> 0 Then
                Percent_change = ((x2 - x1) / x1)
                ws.Cells(summary_table_row, 11).Value = Percent_change
                ws.Cells(summary_table_row, 11).NumberFormat = ".00%"
            Else
                ws.Cells(summary_table_row, 11).Value = "0%"
            End If
         
            'Coloring cells if they are a negative number of not
            If (Percent_change < 0) Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            ElseIf (Percent_change > 0) Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
            'checking to see if next cell is empty
            ElseIf (IsEmpty(ws.Cells(summary_table_row, 11).Value) = True) Then
            
            End If
            
            'Increment row
            summary_table_row = summary_table_row + 1
        
            'resetting variables
            Volume = 0
            x1 = ws.Cells(i + 1, 3).Value
            x2 = 0
        Else
            'summing the total volume of each ticker symbol
            Volume = Volume + Cells(i, 7).Value
        End If

    Next i
Next ws

MsgBox ("Finished")

End Sub
