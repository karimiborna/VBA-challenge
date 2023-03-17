Sub datareset()

For Each ws In Worksheets

    ws.Columns("J:M").Clear
    ws.Columns("O:Q").Clear
    ws.Columns("K:K").Interior.ColorIndex = 0

Next ws

End Sub