Attribute VB_Name = "Module1"
Sub datafiller()

For Each ws In Worksheets


ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"

ws.Range("O1").Value = "Greatest % Change"
ws.Range("O2").Value = "Greatest % Decrease"
ws.Range("O3").Value = "Greatest Total Volume"

Dim LastRow As Long

Dim thisticker As String
thisticker = ws.Cells(2, 1).Value
LastRow = ActiveSheet.UsedRange.Rows.Count

Dim StartRow As Long
StartRow = 2

Dim EndRow As Long

Dim stockopen As Double
stockopen = ws.Cells(2, 3).Value

Dim stockclose As Double

Dim totalvolume As LongLong
totalvolume = 0

Dim yearlychange As Double

Dim percentchange As Double

Dim tablefiller As Long
tablefiller = 2

'loop through each stock
For i = 2 To LastRow
    
    If ws.Cells(i, 1) = thisticker Then
        '*CONTINUAL SUM OF TOTAL VOLUME THROUGH LOOP*
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1) <> thisticker Then
        
        '*SET CLOSE VARIABLES*
            EndRow = i
            stockclose = ws.Cells(i, 6).Value
            
            '*TICKER*
            ws.Cells(tablefiller, 10).Value = ws.Cells(i, 1).Value
            
            '*YEARLY CHANGE*
             yearlychange = stockclose - stockopen
             ws.Cells(tablefiller, 11).Value = yearlychange
             
                If ws.Cells(tablefiller, 11).Value > 0 Then
                    ws.Cells(tablefiller, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(tablefiller, 11).Interior.Color = vbRed
                End If
            
            '*PERCENT CHANGE*
            percentchange = (stockclose / stockopen) - 1
            
            
            ws.Cells(tablefiller, 12).Value = FormatNumber(percentchange)
            ws.Cells(tablefiller, 12).NumberFormat = "0.00"
            'ws.Cells(tablefiller, 12).Style = "Percentage"
                'If biggest increase yet, push ticker to cells and % to next cells
                If percentchange > 0 And percentchange > ws.Range("Q1").Value Then
                    ws.Range("P1").Value = thisticker
                    ws.Range("Q1").Value = percentchange
                End If
                If percentchange < 0 And percentchange < ws.Range("Q2").Value Then
                    ws.Range("P2").Value = thisticker
                    ws.Range("Q2").Value = percentchange
                End If
                
            '*SET TOTAL VOLUME*
            ws.Cells(tablefiller, 13).Value = totalvolume
            If totalvolume > ws.Range("Q3").Value Then
                ws.Range("P3").Value = thisticker
                ws.Range("Q3").Value = totalvolume
            End If
            '*Reset variables that can be found on first row for next loop*
            StartRow = i + 1
            thisticker = ws.Cells(i + 1, 1).Value
            stockopen = ws.Cells(i + 1, 3).Value
            tablefiller = tablefiller + 1
            totalvolume = 0
                
        
        End If
    
    End If
    
Next i

Next ws

MsgBox ("Populate complete.")
End Sub
Sub datareset()

For Each ws In Worksheets

    ws.Columns("J:M").Clear
    ws.Columns("O:Q").Clear
    ws.Columns("K:K").Interior.ColorIndex = 0

Next ws

End Sub
