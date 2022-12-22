# VBA-challenge
Sub stocks()

Dim WS As Worksheet
Dim wb As Workbook
Set wb = ActiveWorkbook


For Each WS In wb.Sheets

'add column names
WS.Range("a1:q1").Value = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Vol", " ", " ", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", " ", " ", "Ticker", "Value")


Dim Ticker_Name As String
Dim open_price As Double
Dim close_price As Double
Dim change As Double
Dim percent As Double
Dim volume As Variant
volume = 0
Dim max_ticker As String
Dim max_percent As Double
Dim min_ticker As String
Dim min_percent As Double
Dim greatest_vol_ticker As String
Dim greatest_vol As Variant
greatest_vol = 0
Dim table As Long
table = 2

lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'initial price of stock
open_price = WS.Cells(2, 3).Value

'Loop
For i = 2 To lastrow

'start volume count
volume = volume + WS.Cells(i, 7).Value


'check ticker name
If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then


'add name to summary table
Ticker_Name = WS.Cells(i, 1).Value
WS.Range("J" & table).Value = Ticker_Name

'take close price
close_price = WS.Cells(i, 6).Value

'what is the differnce
change = close_price - open_price
WS.Range("k" & table).Value = change

'add color
If (change > 0) Then
WS.Range("k" & table).Interior.ColorIndex = 4
ElseIf (change <= 0) Then
WS.Range("k" & table).Interior.ColorIndex = 3
End If

'add percent
percent = (change / open_price) * 100
WS.Range("L" & table).Value = percent & " %"

'add total volume to table
WS.Range("m" & table).Value = volume

'next row
table = table + 1
'reset price
open_price = WS.Cells(i + 1, 3).Value

'add 2nd table
WS.Range("O2").Value = "Greatest % increase"
WS.Range("O3").Value = "Greatest % decrease"
WS.Range("O4").Value = "Greatest Total Volume"


'find greatest and lowest increase
If (percent > max_percent) Then
max_percent = percent
max_ticker = Ticker_Name

ElseIf (percent < min_percent) Then
min_percent = percent
min_ticker = Ticker_Name
End If

'find volume
If (volume > greatest_vol) Then
greatest_vol = volume
WS.Range("Q4").Value = greatest_vol
greatest_vol_ticker = Ticker_Name
End If





'reset
volume = 0
percent = 0

'add to 2nd table

WS.Range("P2").Value = max_ticker
WS.Range("Q2").Value = max_percent & "%"
WS.Range("p3").Value = min_ticker
WS.Range("Q3").Value = min_percent & "%"
WS.Range("p4").Value = greatest_vol_ticker









End If
Next i


Next WS




 End Sub

