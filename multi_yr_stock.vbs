Attribute VB_Name = "Module1"
Sub stocks()

Dim WS As Worksheet
Dim wb As Workbook
Set wb = ActiveWorkbook


For Each WS In wb.Sheets

'add column names
WS.Range("a1:q1").Value = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Vol", " ", " ", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", " ", " ", "Ticker", "Value")


Dim ticker_name As String
ticker_name = " "
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim change As Double
change = 0
Dim percent As Double
percent = 0
Dim volume As Variant
volume = 0
Dim max_ticker As String
max_ticker = 0
Dim max_percent As Double
max_percent = 0
Dim min_ticker As String
min_ticker = 0
Dim min_percent As Double
min_percent = 0
Dim greatest_vol_ticker As String
reatest_vol_ticker = 0
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
ticker_name = WS.Cells(i, 1).Value
WS.Range("J" & table).Value = ticker_name

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


'find greatest and lowest increase
If (percent > max_percent) Then
max_percent = percent
max_ticker = ticker_name

ElseIf (percent < min_percent) Then
min_percent = percent
min_ticker = ticker_name
End If

'find volume
If (volume > greatest_vol) Then
greatest_vol = volume
greatest_vol_ticker = ticker_name
End If


'reset
volume = 0
percent = 0

Else
volume = volume + WS.Cells(i, 7).Value
End If

Next i

'add to 2nd table
WS.Range("Q4").Value = greatest_vol
WS.Range("P2").Value = max_ticker
WS.Range("Q2").Value = max_percent & "%"
WS.Range("p3").Value = min_ticker
WS.Range("Q3").Value = min_percent & "%"
WS.Range("p4").Value = greatest_vol_ticker
WS.Range("o2").Value = "Greatest % increase"
WS.Range("o3").Value = "Greatest % decrease"
WS.Range("o4").Value = "Greatest Total Volume"




Next WS

max_percent = 0
min_percent = 0
greatest_vol = 0


 End Sub


