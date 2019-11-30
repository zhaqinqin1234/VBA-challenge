Sub totaVolume():
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet

    Dim ticker As String
    Dim total As Double
    Dim tickerSummaryRow As Integer
    Dim newTickerRow As Double
    Dim lastRow As Double
    Dim openPrice As Double
    Dim closePrice As Double
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
    tickerSummaryRow = 2
    total = 0
    newTickerRow = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       For i = 2 To lastRow
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
           ticker = ws.Cells(i, 1).Value
           total = total + ws.Cells(i, 7).Value
          
           ws.Cells(tickerSummaryRow, 9).Value = ticker
           ws.Cells(tickerSummaryRow, 12).Value = total
           openPrice = ws.Cells(newTickerRow, 3).Value
           closePrice = ws.Cells(i, 6).Value
           
           ws.Cells(tickerSummaryRow, 10).Value = closePrice - openPrice
           
                If openPrice = 0 Or IsEmpty(ws.Cells(newTickerRow, 3).Value) Then
                    ws.Cells(tickerSummaryRow, 11).Value = "Null"
                Else: ws.Cells(tickerSummaryRow, 11).Value = ws.Cells(tickerSummaryRow, 10).Value / openPrice
                End If
                
            tickerSummaryRow = tickerSummaryRow + 1
            newTickerRow = i + 1
            total = 0
         
        Else
            total = total + ws.Cells(i, 7).Value
                
        End If
Next i
Next 
starting_ws.Activate
End Sub

Sub conditionalFormatting():

For Each ws In Worksheets

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(i, 11).NumberFormat = "0.00%"
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3

    End If
Next i
Next ws
End Sub




