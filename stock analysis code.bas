Attribute VB_Name = "Module1"
Sub multiple_year_stock_data():
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        Call ProcessMultiple_year_stock_data(ws)
        Next ws
    End Sub
Sub ProcessMultiple_year_stock_data(ws As Worksheet)

    ' Define all the neccessary variables
  
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim rowStart As Long
    Dim lastRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestVolume As String
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    totalVolume = 0
    rowStart = 2
    

    ' ^ process data
'begin loop to read each row

For i = 2 To lastRow

'check ticker symbol range
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'assign ticker
ticker = ws.Cells(i, 1).Value

'calculate total volume
totalVolume = totalVolume + ws.Cells(i, 7).Value

'Assign closing price
closePrice = ws.Cells(i, 6).Value


'Calculate yearly change open to close
yearlyChange = closePrice - openPrice

'Calculate percent change
If openPrice <> 0 Then
    percentChange = (yearlyChange / openPrice)
Else
    percentChange = 0
    
End If
'add ticker symbol, yearly change, percent change, & total volume columns

ws.Range("I" & rowStart).Value = ticker
ws.Range("J" & rowStart).Value = yearlyChange
ws.Range("k" & rowStart).Value = percentChange
ws.Range("L" & rowStart).Value = totalVolume

'Check conditional formatting for positive and negative change

If yearlyChange >= 0 Then
    ws.Range("J" & rowStart).Interior.Color = vbGreen
Else
    ws.Range("J" & rowStart).Interior.Color = vbRed
End If
'check and update greatest increase, decrease, & volume
  If percentChange > greatestIncrease Then
    greatestIncrease = percentChange
    tickerGreatestIncrease = ticker

ElseIf percentChange < greatestDecrease Then
    greatestDecrease = percentChange
    tickerGreatestDecrease = ticker
    
 End If
   If totalVolume > greatestVolume Then
    greatestVolume = totalVolume
    tickerGreatestVolume = ticker
 End If
    
    'reset totalvolume for next ticker
    totalVolume = 0
' move to next row
    rowStart = rowStart + 1
' assign new start price for the next ticker
    openPrice = ws.Cells(i + 1, 3).Value
Else
    ' total volume for the current ticker symbol
    totalVolume = totalVolume + ws.Cells(i, 7).Value
End If
            
Next i

'summary table greatest increase, decrease, and volume
ws.Cells(3, 15).Value = tickerGreatestIncrease
ws.Cells(3, 16).Value = greatestIncrease
ws.Cells(4, 15).Value = tickerGreatestDecrease
ws.Cells(4, 16).Value = greatestDecrease
ws.Cells(5, 15).Value = tickerGreatestVolume
ws.Cells(5, 16).Value = greatestVolume

'summary table number format
ws.Cells(3, 16).NumberFormat = "0.00%"
ws.Cells(4, 16).NumberFormat = "0.00%"

'summary table headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
 
ws.Cells(2, 15).Value = "Ticker"
ws.Cells(2, 16).Value = "Value"

ws.Cells(3, 14).Value = "Greatest % Increase"
ws.Cells(4, 14).Value = "Greatest % Decrease"
ws.Cells(5, 14).Value = "Greatest Total Volume"

End Sub
