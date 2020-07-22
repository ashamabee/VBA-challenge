Attribute VB_Name = "Module21"
Sub Stocks2()
Dim ws As Worksheet
For Each ws In Worksheets
Dim Ticker As String
Ticker = ws.Cells(2, 1).Value
Dim OutputRow As Long
OutputRow = 2
Dim SumVolume As Variant
Dim OpenStock As Long
OpenStock = ws.Cells(2, 3).Value
Dim CloseStock As Double
Dim RowCount As Variant
Dim GreatestIncreaseValue As Double
GreatestIncreaseValue = 0
Dim GreatestDecreaseValue As Double
GreatestDecreaseValue = 0
Dim GreatestIncrease As String
Dim GreatestDecrease As String
RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Find RowCount
RowCount = RowCount - 2
    
    ws.Range("I1").Value = "Ticker"        'Headers and font and fit
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("I1:L1").Columns.AutoFit
  'Ticker = ws.Cells(2, 1).Value
    ws.Range("I2").Value = ws.Range("A2").Value
    For i = 2 To RowCount
    
        If Ticker <> ws.Cells(i, 1).Value Then 'if cell = cell below,
            Ticker = ws.Cells(i, 1).Value        'Set variable ticker to the current cell's ticker
            CloseStock = ws.Cells(i, 6).Value
            ws.Cells(OutputRow, 10).Value = CloseStock - OpenStock
            OpenStock = ws.Cells(i, 3).Value
            
            ws.Cells(OutputRow + 1, "I").Value = Ticker
            ws.Cells(OutputRow, "L").Value = SumVolume
            
                If OpenStock > 0 And CloseStock > 0 Then
                    ws.Cells(OutputRow, "K").Value = ws.Cells(OutputRow, "J").Value / OpenStock
                    ws.Cells(OutputRow, "K").NumberFormat = "0.00%"
                    
                        If ws.Cells(OutputRow, "K").Value > ws.Cells(OutputRow + 1, "K").Value Then
                        GreatestIncreaseValue = ws.Cells(OutputRow, "K").Value
                        GreatestIncrease = ws.Cells(OutputRow, "I").Value
                        Else
                        GreatestIncreaseValue = ws.Cells(OutputRow + 1, "K").Value
                        GreatestIncrease = ws.Cells(OutputRow + 1, "I").Value
                        End If
                        
                        If ws.Cells(OutputRow, "K").Value < ws.Cells(OutputRow + 1, "K").Value Then
                        GreatestDecreaseValue = ws.Cells(OutputRow, "K").Value
                        GreatestDecrease = ws.Cells(OutputRow, "I").Value
                        Else
                        GreatestDecreaseValue = ws.Cells(OutputRow + 1, "K").Value
                        GreatestDecrease = ws.Cells(OutputRow + 1, "I").Value
                        End If
                End If
            OutputRow = OutputRow + 1
            SumVolume = 0
        Else
            'Ticker = Cells(i + 1, 1).Value
            SumVolume = SumVolume + ws.Cells(i, 7).Value
        End If
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
                        ws.Range("O2").Value = "Greatest % Increase"
                        ws.Range("O3").Value = "Greatest % Decrease"
                        ws.Range("O4").Value = "Greatest Volume"
                        ws.Range("Q2").Value = GreatestIncreaseValue
                        ws.Range("P2").Value = GreatestIncrease
                        ws.Range("Q3").Value = GreatestDecreaseValue
                        ws.Range("P3").Value = GreatestDecrease
                        ws.Range("Q3").NumberFormat = "0.00%"
                        ws.Range("Q2").NumberFormat = "0.00%"
                        ws.Range("O2:O4").Font.Bold = True
                        ws.Range("O2:O4").Columns.AutoFit
    
Next ws
End Sub
