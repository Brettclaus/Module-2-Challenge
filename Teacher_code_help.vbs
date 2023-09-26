Attribute VB_Name = "Module2"
Option Explicit

'BEGIN AT THE BEGINNING - LEVERAGE THE CREDIT CARD EXAMPLE

'SIMPLY UPDATE VOLUME FIRST

Sub WallStreet_1()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long

' Name headers where values are expected to go
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        vol = 0
        j = j + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub

'ADD PRICE CHANGE INTO THE MIX

Sub WallStreet_2()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long
Dim Delta As Double

' Name headers where values are expected to go
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        'Range("J" & 2 + j).NumberFormat = "0.00"
        
        vol = 0
        Delta = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub

' ADD % CHANGE

Sub WallStreet_3()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long
Dim Delta, percChange As Double

' Name headers where values are expected to go
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        'Range("J" & 2 + j).NumberFormat = "0.00"
        
        percChange = Delta / Cells(openIdx, 3).Value
        Range("K" & 2 + j).Value = percChange
        'Range("K" & 2 + j).NumberFormat = "0.00%"
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub

' COLORIZE CELLS BASED ON VALUE OF DELTA

Sub WallStreet_4()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long
Dim Delta, percChange As Double

' Name headers where values are expected to go
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        Range("J" & 2 + j).NumberFormat = "0.00"
        
        If Delta > 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Delta < 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Delta / Cells(openIdx, 3).Value
        Range("K" & 2 + j).Value = percChange
        'Range("K" & 2 + j).NumberFormat = "0.00%"
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i

End Sub

' ROLL UP THE BROAD METRICS BASED ON THE VALUES COMPUTED ABOVE

Sub WallStreet_5()

Dim i, j, openIdx As Integer
Dim vol, rowCount, rowCountK, rowCountL As Long
Dim Delta, percChange As Double
Dim bestIdx, worstIdx, mostIdx As Integer

' Name headers where values are expected to go
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        'Range("J" & 2 + j).NumberFormat = "0.00"
        
        If Delta > 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Delta < 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Delta / Cells(openIdx, 3).Value
        Range("K" & 2 + j).Value = percChange
        'Range("K" & 2 + j).NumberFormat = "0.00%"
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If
    
Next i

rowCountK = Cells(Rows.Count, "K").End(xlUp).Row
rowCountL = Cells(Rows.Count, "L").End(xlUp).Row

' Compute the stocks that had the biggest % changes in price and volume
Range("Q2") = WorksheetFunction.Max(Range("K2:K" & rowCountK).Value)
Range("Q3") = WorksheetFunction.Min(Range("K2:K" & rowCountK).Value)
'Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCountK))
'Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCountK))
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCountL).Value)

' Use MATCH to find the index of the best, worst and most
bestIdx = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCountK).Value), Range("K2:K" & rowCountK).Value, 0)
worstIdx = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCountK).Value), Range("K2:K" & rowCountK).Value, 0)
mostIdx = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCountL).Value), Range("L2:L" & rowCountL).Value, 0)

' Plug the values into the required cells
Range("P2").Value = Cells(bestIdx + 1, 9).Value
'Range("P2").NumberFormat = "0.00%"

Range("P3").Value = Cells(worstIdx + 1, 9).Value
'Range("P3").NumberFormat = "0.00%"

Range("P4").Value = Cells(mostIdx + 1, 9).Value

End Sub

' UPDATE ALL SHEETS TOGETHER

Sub WallStreet_6()

Dim ws As Object

For Each ws In Worksheets

Dim i, j, openIdx As Integer
Dim vol, rowCount, rowCountK, rowCountL As Long
Dim Delta, percChange As Double
Dim bestIdx, worstIdx, mostIdx As Integer


' Name headers where values are expected to go
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2


rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        vol = vol + ws.Cells(i, 7).Value
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("L" & 2 + j).Value = vol
        
        Delta = ws.Cells(i, 6).Value - ws.Cells(openIdx, 3).Value
        ws.Range("J" & 2 + j).Value = Delta
        'ws.Range("J" & 2 + j).NumberFormat = "0.00"
        
        If Delta > 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Delta < 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Delta / ws.Cells(openIdx, 3).Value
        ws.Range("K" & 2 + j).Value = percChange
        'ws.Range("K" & 2 + j).NumberFormat = "0.00%"
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + ws.Cells(i, 7).Value
        
    End If
    
Next i

rowCountK = ws.Cells(Rows.Count, "K").End(xlUp).Row
rowCountL = ws.Cells(Rows.Count, "L").End(xlUp).Row

' Compute the stocks that had the biggest % changes in price and volume
ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q3").NumberFormat = "0.00%"

'ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value)
'ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value)

' Use MATCH to find the index of the best, worst and most
bestIdx = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
worstIdx = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
mostIdx = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value), ws.Range("L2:L" & rowCountL).Value, 0)

' Change format of column K to reflect percentage values
ws.Range("K2:K" & rowCountK).NumberFormat = "0.00%"

' Plug the values into the required cells
ws.Range("P2").Value = ws.Cells(bestIdx + 1, 9).Value

ws.Range("P3").Value = ws.Cells(worstIdx + 1, 9).Value

ws.Range("P4").Value = ws.Cells(mostIdx + 1, 9).Value


Next ws

End Sub


