Attribute VB_Name = "Module1"
Sub YearlyStock()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("J1:M1").EntireColumn.AutoFit

    Dim ticker As String
    Dim sk_change, openvalue, close_price, sk_volume, percent_change As Double
    Dim i, lsrow, Perc_lsrow As Long
    Dim summaryrow As Integer
    
    'sk_change = yearly change, sk_volume = stock volume

    summaryrow = 2
    
    lsrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Perc_lsrow = ws.Cells(Rows.Count, 12).End(xlUp).Row

    openvalue = ws.Range("C2").Value 'open value

    For i = 2 To lsrow

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value  'ticker symbol

            close_price = ws.Cells(i, 6).Value

'calculations
            sk_change = close_price - openvalue
            percent_change = (sk_change / openvalue)
            sk_volume = sk_volume + ws.Cells(i, 7).Value
'fill in summary table
            ws.Range("J" & summaryrow).Value = ticker
            ws.Range("K" & summaryrow).Value = sk_change
            ws.Range("L" & summaryrow).Value = percent_change
            ws.Range("M" & summaryrow).Value = sk_volume
'add to the next row
            summaryrow = summaryrow + 1
'reset the volume to 0 for the next ticker
            sk_volume = 0
'goes to the next open 1st open value
            openvalue = ws.Cells(i + 1, 3).Value
'if following ticker is the same
            Else
                sk_volume = sk_volume + ws.Cells(i, 7).Value
                
            End If
        
    Next i

ws.Range("L2:L" & Perc_lsrow).NumberFormat = "0.00%"
    
Next ws


Call changes

Call Greatest_Stock


End Sub

Sub changes()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim i As Long

    lsrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Will format pos changes green and neg changes red
    For i = 2 To lsrow
            If ws.Cells(i, 11) > 0 Then
            ws.Cells(i, 11).Interior.Color = vbGreen
            ElseIf ws.Cells(i, 11) = 0 Then
            ws.Cells(i, 11).Interior.Color = xlNone
            Else
            ws.Cells(i, 11).Interior.Color = vbRed
            
            
        End If
    
    Next i
Next ws

End Sub

Sub Greatest_Stock()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O1:P1").EntireColumn.AutoFit

    Dim i, summaryrow As Integer
    Dim inc, dec, highvol  As Double


    Perc_lsrow = ws.Cells(Rows.Count, 12).End(xlUp).Row     'last row of percent change column
    Totvol_lsrow = ws.Cells(Rows.Count, 13).End(xlUp).Row   'last row of total volume column

'Greatest % increase
    For i = 2 To Perc_lsrow
            If ws.Cells(i, 12).Value = WorksheetFunction.Max((ws.Range("L1:L" & Perc_lsrow).Value)) Then
                ws.Range("Q2").Value = ws.Cells(i, 12).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = ws.Cells(i, 10).Value
            End If
'Greatest % decrease
            If ws.Cells(i, 12).Value = WorksheetFunction.Min((ws.Range("L1:L" & Perc_lsrow).Value)) Then
                ws.Range("Q3").Value = ws.Cells(i, 12).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = ws.Cells(i, 10).Value
            End If
        Next i
        
'Greatest total volume
    For i = 2 To Totvol_lsrow
            If ws.Cells(i, 13).Value = WorksheetFunction.Max((ws.Range("M1:L" & Totvol_lsrow).Value)) Then
                ws.Range("Q4").Value = ws.Cells(i, 13).Value
                ws.Range("P4").Value = ws.Cells(i, 10).Value
            End If
    Next i


Next ws

End Sub


