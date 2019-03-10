Sub stockmarkettest()

    

    Dim newrowcount As Double
    Dim totalvolume As Double
    
    'loop through all worksheets
    
    For Each ws In Worksheets
    
    'define variables
    newrowcount = 2
    totalvolume = ws.Range("G2").Value
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'create new rows in the worksheet
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "totalvolume"
    
    'loop through all ticker symbols
    
    For i = 2 To lastrow
    
    'check if we're still in the same ticker symbol, if not...
    
        If ws.cells(i + 1, 1) <> ws.cells(i, 1) Then
        ws.Range("I" & newrowcount).Value = ws.Cells(i, 1).Value
        ws.Range("J" & newrowcount).Value = totalvolume
        totalvolume = ws.Range("G" & i + 1)
        newrowcount = newrowcount + 1
        
        Else
        totalvolume = totalvolume + ws.Range("G" & i).Value
        
        
        End If
    Next i
    Next ws
    
End Sub


