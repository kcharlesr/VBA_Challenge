Attribute VB_Name = "Module1"
Sub Stock_Ticker_Exercise()

'Define Sheets and Set up Loop for all Sheets

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

ws.Activate


''Define Last Row

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Define Variables and Set Initial Values

Dim Ticker As String

Dim Open_Price As Double

Dim Close_Price As Double

Dim Yearly_Change As Double

Dim Percent_Change As Double

Dim Volume As Double

Dim Row As Double

Dim Column As Double

Volume = 0

Row = 2

Column = 1

Open_Price = Cells(2, Column + 2).Value


' Create Headers for Summary Table

Cells(1, "I").Value = "Ticker"

Cells(1, "J").Value = "Yearly Change"

Cells(1, "K").Value = "Percent Change"

Cells(1, "L").Value = "Total Stock Volume"


' Create Loop through Tickers until two consecutive tickers no longer match and send ticker name to summary table

For i = 2 To LastRow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Ticker_Name = Cells(i, 1).Value

Cells(Row, 9).Value = Ticker_Name

' Define Close_Price

Close_Price = Cells(i, 6).Value

' Define Yearly_Change

Yearly_Change = Close_Price - Open_Price

Cells(Row, 10).Value = Yearly_Change

' Define Percent_Change

If (Open_Price = 0 And Close_Price = 0) Then

Percent_Change = 0

ElseIf (Open_Price = 0 And Close_Price <> 0) Then

Percent_Change = 1

Else

Percent_Change = Yearly_Change / Open_Price

Cells(Row, 11).Value = Percent_Change

Cells(Row, 11).NumberFormat = "0.00%"

End If

' Define Volume

Volume = Volume + Cells(i, 7).Value

Cells(Row, 12).Value = Volume


'Move down Row

Row = Row + 1

' Reset Open_Price

Open_Price = Cells(i + 1, 3)

' Reset the Volume

Volume = 0

Else

Volume = Volume + Cells(i, 7).Value

End If

Next i



' Define Last Row

YCLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row


' Define Cell Color Format

For j = 2 To YCLastRow

If Cells(j, 10).Value >= 0 Then

Cells(j, 10).Interior.Color = vbGreen

ElseIf Cells(j, 10).Value < 0 Then

Cells(j, 10).Interior.Color = vbRed

End If

Next j



' Create Headers for Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume, Ticker, and Value

Cells(2, Column + 14).Value = "Greatest % Increase"

Cells(3, Column + 14).Value = "Greatest % Decrease"

Cells(4, Column + 14).Value = "Greatest Total Volume"

Cells(1, Column + 15).Value = "Ticker"

Cells(1, Column + 16).Value = "Value"


' Create Loop for Greatest Values and Coresponding Ticker

For r = 2 To YCLastRow

If Cells(r, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & YCLastRow)) Then

Cells(2, 16).Value = Cells(r, 9).Value

Cells(2, 17).Value = Cells(r, 11).Value

Cells(2, 17).NumberFormat = "0.00%"

ElseIf Cells(r, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & YCLastRow)) Then

Cells(3, 16).Value = Cells(r, 9).Value

Cells(3, 17).Value = Cells(r, 11).Value

Cells(3, 17).NumberFormat = "0.00%"

ElseIf Cells(r, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & YCLastRow)) Then

Cells(4, 16).Value = Cells(r, 9).Value

Cells(4, 17).Value = Cells(r, 12).Value

End If

Next r

'Format Column Width to fit Values

Columns("I:Q").AutoFit

Next ws



End Sub



