Sub Stock_Market_Analysis()

'Run on every worksheet
 For Each ws In ActiveWorkbook.Worksheets
 ws.Activate

' Headers
Range("A1").Value = "<ticker>"
Range("B1").Value = "<date>"
Range("C1").Value = "<open>"
Range("D1").Value = "<high>"
Range("E1").Value = "<low>"
Range("F1").Value = "<close>"
Range("G1").Value = "<vol>"
Range("i1").Value = "Ticker Code"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Volume"
Range("i:L").ColumnWidth = 16

' Input N column values

Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("N:N").ColumnWidth = 18

' Set VBA values

Dim TickerName As String
Dim OpenStock As Double
Dim CloseStock As Double
Dim TickerCode As String
Dim Percent_Change As Double
Dim Yearly_Change As Double
Dim TotalStockVolume As Double
Dim xIncrease As Double
Dim xDecrease As Double
Dim GreatestTotalVolume As Double
Dim Row As Double
Dim x As Double

' Set Row

Row = Cells(Rows.Count, "A").End(xlUp).Row
x = 2

' Set opening stock price

OpenStock = Cells(2, 3).Value

For i = 2 To Row
        TickerName = Cells(i, 1).Value
        TickerCode = Cells(i + 1, 1).Value
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    
    If TickerName <> TickerCode Then
        Cells(x, "I").Value = TickerName
        Cells(x, "L").Value = TotalStockVolume
        CloseStock = Cells(i, 6).Value
        
        Yearly_Change = CloseStock - OpenStock
        Cells(x, 10).Value = Yearly_Change
     
' Conditional Formatting

    If Yearly_Change > 0 Then
        Cells(x, 10).Interior.ColorIndex = 4
    
        ElseIf Yearly_Change < 0 Then
            Cells(x, 10).Interior.ColorIndex = 3
    
        Else
            Cells(x, 10).Interior.ColorIndex = 0

End If

' Calculating Percent Change

If (OpenStock = 0 And CloseStock = 0) Then
        Percent_Change = 0
    
        ElseIf (OpenStock = 0 And CloseStock <> 0) Then
            Percent_Change = 1
    
        Else
            Percent_Change = Yearly_Change / OpenStock
            Cells(x, 11).Value = Percent_Change
       
End If
    
' Calculating Total Stock Volume

TotalStockVolume = 0
OpenStock = Cells(i + 1, 3).Value
x = x + 1
        
End If

' Number Formatting
Cells(x, 11).NumberFormat = "0.00%"

Next i

Next ws

End Sub
