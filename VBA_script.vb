Sub calc()

'Looping through all worksheets in the workbook.
For Each ws In Worksheets

    
    
    'Storing the ticker symbol in a variable.
    Dim ticker As String
    ticker = "fudge"
    
    
    'Storing the opening price
    Dim openP As Double
    openP = 0
    
    
    'Closing price
    Dim closeP As Double
    closeP = 0
    
    'Variable to keep track of volume
    Dim volume As Double
    
    'Variables for tracking the opening and closing day, in case data is not chronological
    Dim openDay As Long
    Dim closeDay As Long
    
    'Variables for final calculation of yearly price change and percent change.
    Dim yearChange As Double
    Dim percentChange As Double
    
   
    
    'Figuring out the length of the sheet. Code courtesy of H. Bree.
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Index j for output rows
    Dim j As Long
    j = 1
    
    'Diagnostic MsgBox
    'MsgBox lastRow
    
    ' Writing the header row
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
        
    'Loop grabs the ticker symbols, change, percent change, and volume. Then prints the info.
    For i = 2 To lastRow + 1
        
        
        'This happens when a new ticker symbol is hit.
        If ticker <> ws.Cells(i, 1) Then
            
            'This is actually the last thing to happen for each ticker.
            If j <> 1 Then 'nothing happens for the first ticker, j=1.
                
                ws.Cells(j, 9) = ticker
                
                ws.Cells(j, 10) = yearChange
                
                'yearChange Cell is red if negative, green if positive.
                If yearChange < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
                
                If yearChange > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                End If
                '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            
                'Accounting for some opening prices of zero (nonsense) in worksheet P
                If openP = 0 Then
                    percentChange = 0 'Nonsensical due to nonsensical data
                Else
                    percentChange = yearChange / openP 'yearChange was tracked below.
                End If
                
                'Print and format percent change as percentage.
                ws.Cells(j, 11) = percentChange
                ws.Cells(j, 11).NumberFormat = "0.00%"
                
                'Print Volume
                ws.Cells(j, 12) = volume
            End If
            '----------------------------------------------------------
        
        
            ticker = ws.Cells(i, 1)
            
            
            
            openP = ws.Cells(i, 3)
            closeP = ws.Cells(i, 6)
            
            volume = ws.Cells(i, 7)
            
            'Opening and closing day obviously are not the same. This is just a starting point
            openDay = ws.Cells(i, 2)
            closeDay = ws.Cells(i, 2)
            
            j = j + 1
            
            
        'Here is where building total volume and tracking opening/closing price occurs
        Else
            volume = volume + ws.Cells(i, 7)
            
            'Next two if statements hopefully would account for data out of chronological order.
            'Not that that is necessary with the sheet I am working with.
            If openDay > ws.Cells(i, 2) Then
                openP = ws.Cells(i, 3)
                openDay = ws.Cells(i, 2)
            End If
            
            If closeDay < ws.Cells(i, 2) Then
                closeP = ws.Cells(i, 6)
                closeDay = ws.Cells(i, 2)
            End If
            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
            
            yearChange = closeP - openP
            
        End If
        
    Next i
    
    
    'Declaring variables to track max gain, max loss, and max volume.
    Dim maxPercent As Double
        maxPercent = 0
    Dim minPercent As Double
        minPercent = 0
    Dim maxVolume As Double
        maxVolume = 0
    Dim maxTick As String
    Dim minTick As String
    Dim maxVolTick As String
    'MsgBox j
    
    'Sorting through the first table I generated to find max gain, max loss, and max volume among the stocks on the page.
    For k = 2 To j
    
        If ws.Cells(k, 11) > maxPercent Then
            maxPercent = ws.Cells(k, 11)
            maxTicker = ws.Cells(k, 9)
        End If
        
        If ws.Cells(k, 11) < minPercent Then
            minPercent = ws.Cells(k, 11)
            minTicker = ws.Cells(k, 9)
        End If
        
        If ws.Cells(k, 12) > maxVolume Then
            maxVolume = ws.Cells(k, 12)
            maxVolTick = ws.Cells(k, 9)
        End If
        
        'Checking that last ticker is included.
        'If k = j Then
        '    MsgBox ws.Cells(k, 9) & "is last ticker"
        'End If
        '
        'If k = j - 1 Then
        '    MsgBox ws.Cells(k, 9) & "is last ticker"
        'End If
            
    Next k
    
    'MsgBox k
    
    'Printing max gain, max decrease, and max volume, and labels
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    ws.Range("O2") = minTicker
    ws.Range("O3") = maxTicker
    ws.Range("O4") = maxVolTick
        
    ws.Range("O2") = maxTicker
    ws.Range("P2") = maxPercent
    ws.Range("O3") = minTicker
    ws.Range("P3") = minPercent
    ws.Range("O4") = maxVolTick
    ws.Range("P4") = maxVolume
        
    ws.Range("P2:P3").NumberFormat = "0.00%"
    'MsgBox volume
    'MsgBox j
     

        
    'MsgBox ticker

Next ws
    
    
    
    
End Sub
