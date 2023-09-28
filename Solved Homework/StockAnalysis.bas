Attribute VB_Name = "Module1"
Sub stockSummary()
Dim row, totalVol, outRow, greatVol As Double
Dim yrStart, yrEnd, delta, greatDwn, greatUp, pctChange As Double
Dim tick, upTic, dwnTic, volTic As String
For Each ws In Worksheets
    greatVol = 0
    volTic = ""
    greatDwn = 0
    dwnTic = ""
    greatUp = 0
    upTic = ""
    row = 2 'Start row
    tick = ws.Cells(row, 1) 'first ticker
    totalVol = 0
    outRow = 2 'first output row
    yearStart = ws.Cells(row, 3)
    '''Create summary table headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yr. Change"
    ws.Range("K1") = "Pcnt Change"
    ws.Range("L1") = "Total Volume"
    
    Do While ws.Cells(row, 1) <> 0
        If tick <> ws.Cells(row, 1) Then
            '''add summary info and formatting
            yearEnd = ws.Cells(row - 1, 6)
            
            ws.Cells(outRow, 9) = tick
            
            delta = yearEnd - yearStart
            If delta >= 0 Then
                ws.Cells(outRow, 10).Interior.ColorIndex = 4
                ws.Cells(outRow, 10) = delta
            Else
                ws.Cells(outRow, 10).Interior.ColorIndex = 3
                ws.Cells(outRow, 10) = delta
            End If
           
            pctChange = (yearEnd - yearStart) / yearStart
            ws.Cells(outRow, 11) = (yearEnd - yearStart) / yearStart
            ws.Cells(outRow, 11).NumberFormat = "0.00%"
            ws.Cells(outRow, 12) = totalVol
            '''end summary info
            '''check for biggest delta and volume
            If pctChange > greatUp Then
                greatUp = pctChange
                upTic = tick
            End If
            If pctChange < greatDwn Then
                greatDwn = pctChange
                dwnTic = tick
            End If
            If totalVol > greatVol Then
                greatVol = totalVol
                volTic = tick
            End If
            '''end change
            '''reset variables and increment output row
            tick = ws.Cells(row, 1)
            yearStart = ws.Cells(row, 3)
            totalVol = 0
            
            outRow = outRow + 1
        End If
        
        totalVol = totalVol + ws.Cells(row, 7)
        row = row + 1
    Loop
    
    ws.Range("N2") = "Biggest Drop"
    ws.Range("N3") = "Biggest Rise"
    ws.Range("N4") = "Biggest Volume"
    
    ws.Range("O1") = "Ticker"
    ws.Range("O2") = dwnTic
    ws.Range("O3") = upTic
    ws.Range("O4") = dwnTic
    
    ws.Range("P1") = "Value"
    ws.Range("P2") = greatDwn
    ws.Range("P3") = greatUp
    ws.Range("P4") = greatVol
Next
End Sub
