Sub Stocks()

Dim YearlyChange As Double
Dim PercentChange As Double
Dim TSVolume As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim SummaryRow As Integer
Dim Ticker As String

'For loop for each ws in file
For Each ws In Worksheets

'Establish SummaryRow; keep track of where each value is in table
SummaryRow = 2

'Determine Opening_Price Column
OpeningPrice = ws.Cells(2, 3).Value

'Starting TotalStockVolume is 0
TSVolume = 0

'Add Ticker, YearlyChange, PercentChange, TotalStockVolume as headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Determine Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all stocks
For i = 2 To LastRow

'Check to see if cells are different from one another
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Set Ticker
Ticker = ws.Cells(i, 1).Value

'Print Ticker value to Summary Table
ws.Range("I" & SummaryRow).Value = Ticker

'Calculate YearlyChange
YearlyChange = ws.Cells(i, 6).Value - OpeningPrice

'Print YearlyChange value to Summary Table
ws.Range("J" & SummaryRow).Value = YearlyChange

'Calculate PercentChange
If OpeningPrice = 0 Then
PercentChange = 0
Else: PercentChange = YearlyChange / OpeningPrice

End If

'Print PercentChange value to Summary Table
ws.Range("K" & SummaryRow).Value = PercentChange

'Calculate TSVolume
TSVolume = TSVolume + ws.Cells(i, 7).Value

'Print TSVolume value to Summary Table
ws.Range("L" & SummaryRow).Value = TSVolume

'Add one to Summary Row
SummaryRow = SummaryRow + 1

'Reset TSVolume to 0
TSVolume = 0

'Set OpeningPrice to next ticker
OpeningPrice = ws.Cells(i + 1, 3).Value

'If the cell following the row is the same Ticker
Else: TSVolume = TSVolume + ws.Cells(i, 7).Value

End If
Next i

'Add Functionality; GreatestPerInc, GreatestPerDec, GreatestTSVolume
'Add Ticker, Value, and Greatest___ Table
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest Percent Increase"
ws.Range("O3").Value = "Greatest Percent Decrease"
ws.Range("O4").Value = "Greatest Total Stock Volume"

'Find Greatest Percent Max Ticker and Value
For i = 2 To LastRow
If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) Then
'Add max Ticker Value to table
ws.Range("P2").Value = ws.Cells(i, 9).Value
'Add max Value to table
ws.Range("Q2").Value = ws.Cells(i, 11).Value

End If
Next i

'Find Greatest Percent Min Ticker and Value
For i = 2 To LastRow
If ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) Then
'Add min Ticker Value to table
ws.Range("P3").Value = ws.Cells(i, 9).Value
'Add min Value to table
ws.Range("Q3").Value = ws.Cells(i, 11).Value

End If
Next i

'Find Greatest Total Stock Volume Ticker and Value
For i = 2 To LastRow
If ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) Then
'Add max Greatest Total Stock Value to table
ws.Range("P4").Value = ws.Cells(i, 9).Value
'Add max Value to table
ws.Range("Q4").Value = ws.Cells(i, 12).Value

End If
Next i

'Conditional Formatting
'Check if YearlyChange is positive
For i = 2 To LastRow
If ws.Cells(i, 10) > 0 Then
'Change cells to green
ws.Cells(i, 10).Interior.ColorIndex = 4
'If YearlyChange is negative
ElseIf ws.Cells(i, 10) < 0 Then
'Change cells to red
ws.Cells(i, 10).Interior.ColorIndex = 3

End If
Next i

'Change PercentChange column to %
For i = 2 To LastRow
ws.Cells(i, 11).NumberFormat = "0.00%"

Next i
Next ws
End Sub





