Attribute VB_Name = "Module1"
Sub stockBoi()
Dim ticker As Variant
Dim tickCounter As Integer
Dim dateCounter As Variant
Dim stockVolume As Variant
Dim stockCounter As Integer
Dim largeInc As Variant
Dim smallInc As Variant
Dim largeTick As Variant
Dim smallTick As Variant
Dim volTick As Variant
Dim volInc As Variant
Dim dateHolder As Variant
Dim opening As Variant
Dim closing As Variant
tickCounter = 3
dateCounter = 2
stockVolume = 0
stockCounter = 2
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    dateHolder = ws.Name
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ticker = ws.Cells(i, 1).Value
        If ticker = ws.Cells(i + 1, 1).Value Then
            stockVolume = ws.Cells(i, 7).Value + stockVolume
        ElseIf ticker <> ws.Cells(i + 1, 1).Value Then
            stockVolume = ws.Cells(i, 7).Value + stockVolume
            ws.Cells(stockCounter, 12).Value = stockVolume
            stockVolume = 0
            stockCounter = stockCounter + 1
        End If
        If ticker = ws.Cells(2, 1).Value Then
                ws.Cells(2, 9).Value = ticker
        ElseIf ticker <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(tickCounter, 9).Value = ticker
                tickCounter = tickCounter + 1
        End If
    Next i
    For j = 2 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        Dim percent As Variant
        ticker = ws.Cells(j, 1).Value
        percent = 0
        If ws.Cells(j, 2).Value = dateHolder + "0102" Then
            opening = ws.Cells(j, 3).Value
        ElseIf ws.Cells(j, 2).Value = dateHolder + "1231" Then
            closing = ws.Cells(j, 6).Value
            ws.Cells(dateCounter, 10).Value = closing - opening
            ws.Cells(dateCounter, 11).Value = -1 * (1 - closing / opening)
            ws.Cells(dateCounter, 11).NumberFormat = "0.00%"
            If ws.Cells(dateCounter, 10).Value < 0 Then
                ws.Cells(dateCounter, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(dateCounter, 10).Value > 0 Then
                ws.Cells(dateCounter, 10).Interior.ColorIndex = 4
            End If
             dateCounter = dateCounter + 1
        End If
    Next j
    largeInc = ws.Cells(2, 11).Value
    smallInc = ws.Cells(2, 11).Value
    volInc = ws.Cells(2, 12).Value
    For r = 2 To ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        If ws.Cells(r, 11).Value > largeInc Then
            largeInc = ws.Cells(r, 11).Value
            largeTick = ws.Cells(r, 9).Value
        End If
        If ws.Cells(r, 11).Value < smallInc Then
            smallInc = ws.Cells(r, 11).Value
            smallTick = ws.Cells(r, 9).Value
        End If
        If ws.Cells(r, 12).Value > volInc Then
            volInc = ws.Cells(r, 12).Value
            volTick = ws.Cells(r, 9).Value
        End If
    Next r
    ws.Range("Q2").Value = largeInc
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = largeTick
    ws.Range("Q3").Value = smallInc
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = smallTick
    ws.Range("P4").Value = volTick
    ws.Range("Q4").Value = volInc
    tickCounter = 3
    dateCounter = 2
    stockVolume = 0
    stockCounter = 2
Next ws
End Sub


