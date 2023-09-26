Attribute VB_Name = "Module1"
Sub StockCalc()

    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim i, rowCounter, openIdx As Integer
        Dim totalVolume, rowCount, rowCountK, rowCountL As Long
        Dim yearlyChange, percentChange As Double
        Dim bestIndex, worstIndex, mostIndex As Integer

        ' Set column headers for data analysis
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker with Max Change"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Ticker with Most Volume Change"

        ' Initialize variables for calculations
        totalVolume = 0
        rowCounter = 0
        openIdx = 2

        ' Find the last row with data in column A
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

        ' Loop through each row of data
        For i = 2 To rowCount

            ' Check if the next row has a different ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Calculate and populate total stock volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ws.Range("I" & 2 + rowCounter).Value = ws.Cells(i, 1).Value
                ws.Range("L" & 2 + rowCounter).Value = totalVolume

                ' Calculate and populate yearly change
                yearlyChange = ws.Cells(i, 6).Value - ws.Cells(openIdx, 3).Value
                ws.Range("J" & 2 + rowCounter).Value = yearlyChange

                ' Apply conditional formatting for yearly change
                With ws.Range("J" & 2 + rowCounter)
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(1).Interior.ColorIndex = 3 ' Red for negative change
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    .FormatConditions(2).Interior.ColorIndex = 4 ' Green for positive change
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
                    .FormatConditions(3).Interior.ColorIndex = 0 ' No color for zero change
                End With

                ' Calculate and populate percent change
                percentChange = yearlyChange / ws.Cells(openIdx, 3).Value
                ws.Range("K" & 2 + rowCounter).Value = percentChange
                ws.Range("K" & 2 + rowCounter).NumberFormat = "0.00%"

                ' Apply conditional formatting for percent change
                With ws.Range("K" & 2 + rowCounter)
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(1).Interior.ColorIndex = 3 ' Red for negative percent change
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                    .FormatConditions(2).Interior.ColorIndex = 4 ' Green for positive percent change
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
                    .FormatConditions(3).Interior.ColorIndex = 0 ' No color for zero percent change
                End With

                ' Reset variables for the next ticker symbol
                totalVolume = 0
                yearlyChange = 0
                percentChange = 0
                rowCounter = rowCounter + 1
                openIdx = openIdx + 1
            Else
                ' Accumulate volume for the same ticker symbol
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Find the last rows with data in columns K and L
        rowCountK = ws.Cells(Rows.Count, "K").End(xlUp).Row
        rowCountL = ws.Cells(Rows.Count, "L").End(xlUp).Row

        ' Compute and display the stocks with the biggest % changes and most volume change
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value)
        ws.Range("Q2").NumberFormat = "0.00%"

        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value)
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value)

        ' Use MATCH to find the index of the best, worst, and most volume changes
        bestIndex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
        worstIndex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
        mostIndex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value), ws.Range("L2:L" & rowCountL).Value, 0)

        ' Change format of column K to reflect percentage values
        ws.Range("K2:K" & rowCountK).NumberFormat = "0.00%"

        ' Populate the cells with the values of the best, worst, and most volume changes
        ws.Range("P2").Value = ws.Cells(bestIndex + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(worstIndex + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(mostIndex + 1, 9).Value
    Next ws
End Sub


