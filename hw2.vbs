Sub Easy()
Dim LastRow As Double
Dim TickerCount As Integer
Dim i As Double
Dim TickerNow As String
Dim TotalVol As Double
Dim YearEnd As Double
Dim YearStart As Double
Dim PercentChange As Double
Dim BigTicker As String
Dim BigInc As Double
Dim BigDec As Double


'Loop through all of the sheets
For Each ws In Worksheets
    
'Stores the last row number in each sheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create a ticker counter start at 2 bc we have a header
TickerCount = 2

'Create new headers for my calculations
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


'Get the first year start
YearStart = ws.Cells(2, 3).Value

For i = 2 To LastRow
Dim VolNow As Double
VolNow = ws.Cells(i, 7).Value
TotalVol = TotalVol + VolNow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerNow = ws.Cells(i, 1).Value
        ws.Cells(TickerCount, 9).Value = TickerNow
        
        'get the year end value, and condition the cells with if statement
        YearEnd = ws.Cells(i, 6).Value
        ws.Cells(TickerCount, 10).Value = YearEnd - YearStart
        If ws.Cells(TickerCount, 10).Value > 0 Then
            ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
        End If
        
        'calculate percent change and format cells
        If YearStart > 0 Then
            PercentChange = (YearEnd - YearStart) / YearStart
            ws.Cells(TickerCount, 11).Value = PercentChange
            ws.Cells(TickerCount, 11).NumberFormat = "0.00%"
            'get the next start value
            YearStart = ws.Cells(i + 1, 3).Value
        Else
            YearStart = ws.Cells(i + 1, 3).Value
            PercentChange = (YearEnd - YearStart) / YearStart
            ws.Cells(TickerCount, 11).Value = PercentChange
            ws.Cells(TickerCount, 11).NumberFormat = "0.00%"
        End If
        
        
        'record the total volume
        ws.Cells(TickerCount, 12).Value = TotalVol
        TickerCount = TickerCount + 1
        
        'Restart total volume with the next cell
        TotalVol = ws.Cells(TickerCount, 7).Value
    End If
    
  
Next i

'create cell labels
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


Next ws

End Sub