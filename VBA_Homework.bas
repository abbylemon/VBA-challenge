Attribute VB_Name = "Module1"
Sub summary_table()

    'create a string to store all of the ticker types
    Dim ticker As String
    
    'create a variable that will tell us what line to store the results in
    Dim line As Integer
    line = 2
    
    'create a variable to count what row we are looping through
    Dim count As Integer
    count = 0

    'create variables to store the open and close values and variables to do math with those variables
    Dim o, c As Double
    o = 0
    c = 0
    Dim yearlychange As Double
    yearlychange = 0

    'create variables to store the high and low values and variables to do math with those variables
    Dim high, low As Double
    high = 0
    low = 0
    Dim percentchange As Double
    percentchange = 0
    
    'create a variable to store the total volumn for each ticker
    Dim vol As Double
    vol = 0
    
    'find the last row of data
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells(Rows.count, "A").End(xlUp).Row
    
    'print headers for the new table
    Range("I" & 1).Value = "Ticker"
    Range("J" & 1).Value = "Yearly Charge"
    Range("K" & 1).Value = "Percent Change"
    Range("L" & 1).Value = "Total Stock Volume"
    
        'loop though all of the rows in the tab ignoring the header
        For i = 2 To LastRow
        
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                
                o = Cells(i, 3).Value
        
            'check for when the next ticker is different from this one
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'then store that value in the ticker string and print it in the new table
                ticker = Cells(i, 1).Value
                Range("I" & line).Value = ticker
                
                'store the close value for that ticker and print the yearly change
                c = Cells(i, 6).Value
                yearlychange = c - o
                Range("J" & line).Value = yearlychange
                
                'edit the colors of the yearlychange
                If yearlychange < 0 Then
                    Cells(line, 10).Interior.ColorIndex = 3
                Else
                    Cells(line, 10).Interior.ColorIndex = 4
                End If
                
                'print the percent change and format it to a percent
                If c = 0 Then
                    yearlychange = 0
                    Range("K" & line).Value = 0
                    Range("K" & line).NumberFormat = "0.00%"
                    
                Else
                    Range("K" & line).Value = (yearlychange / o)
                    Range("K" & line).NumberFormat = "0.00%"
                    
                End If
                
                'add the final volumn value and print it
                vol = vol + Cells(i, 7).Value
                Range("L" & line).Value = vol
                
                'move on to the next line for the new table
                line = line + 1
                
                'clear the variables
                count = 0
                c = 0
                o = 0
                high = 0
                low = 0
                vol = 0
                
            'for when the tickers are the same
            Else
                
                vol = vol + Cells(i, 7).Value
                
            End If
            
        Next i
        

End Sub
Sub summary_table_challenge()

'CHALLENGE loop through all of the worksheets
For Each ws In Worksheets

    'create a string to store all of the ticker types
    Dim ticker As String
    
    'create a variable that will tell us what line to store the results in
    Dim line As Integer
    line = 2
    
    'create a variable to count what row we are looping through
    Dim count As Integer
    count = 0

    'create variables to store the open and close values and variables to do math with those variables
    Dim o, c As Double
    o = 0
    c = 0
    Dim yearlychange As Double
    yearlychange = 0

    'create variables to store the high and low values and variables to do math with those variables
    Dim high, low As Double
    high = 0
    low = 0
    Dim percentchange As Double
    percentchange = 0
    
    'create a variable to store the total volumn for each ticker
    Dim vol As Double
    vol = 0
    
    'find the last row of data
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.count, "A").End(xlUp).Row
    
    'print headers for the new table
    ws.Range("I" & 1).Value = "Ticker"
    ws.Range("J" & 1).Value = "Yearly Charge"
    ws.Range("K" & 1).Value = "Percent Change"
    ws.Range("L" & 1).Value = "Total Stock Volume"
    
        'loop though all of the rows in the tab ignoring the header
        For i = 2 To LastRow
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                o = ws.Cells(i, 3).Value
        
            'check for when the next ticker is different from this one
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'then store that value in the ticker string and print it in the new table
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & line).Value = ticker
                
                'store the close value for that ticker and print the yearly change
                c = ws.Cells(i, 6).Value
                yearlychange = c - o
                ws.Range("J" & line).Value = yearlychange
                
                'edit the colors of the yearlychange
                If yearlychange < 0 Then
                    ws.Cells(line, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(line, 10).Interior.ColorIndex = 4
                End If
                
                'print the percent change and format it to a percent
                If o = 0 Then
                    yearlychange = 0
                    ws.Range("K" & line).Value = 0
                    ws.Range("K" & line).NumberFormat = "0.00%"
                    
                Else
                    ws.Range("K" & line).Value = (yearlychange / o)
                    ws.Range("K" & line).NumberFormat = "0.00%"
                    
                End If
                
                'add the final volumn value and print it
                vol = vol + ws.Cells(i, 7).Value
                ws.Range("L" & line).Value = vol
                
                'move on to the next line for the new table
                line = line + 1
                
                'clear the variables
                count = 0
                c = 0
                o = 0
                high = 0
                low = 0
                vol = 0
                
            'for when the tickers are the same
            Else
                
                vol = vol + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'challenge declare new variables and cell names
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        Dim percentMin As Double
        Dim percentMinTicker As String
        Dim percentMax As Double
        Dim percentMaxTicker As String
        Dim VolMax As Double
        Dim VolMaxTicker As String
        
        LastRowSummary = ws.Cells(Rows.count, "I").End(xlUp).Row
        
        'challenge find max and min for percent and vol
        percentMax = 0
        percentMin = 0
        VolMax = 0
        For j = 2 To LastRowSummary
            If ws.Cells(j, 11).Value > percentMax Then
                percentMax = ws.Cells(j, 11).Value
                ws.Range("Q2").Value = percentMax
                percentMaxTicker = ws.Cells(j, 9).Value
                ws.Range("P2").Value = percentMaxTicker
                ws.Range("Q2").Value = percentMax
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(j, 11).Value < percentMin Then
                percentMin = ws.Cells(j, 11).Value
                percentMinTicker = ws.Cells(j, 9).Value
                ws.Range("Q3").Value = percentMin
                ws.Range("P3").Value = percentMinTicker
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(j, 12).Value > VolMax Then
                VolMax = ws.Cells(j, 12).Value
                VolMaxTicker = ws.Cells(j, 9).Value
                ws.Range("Q4").Value = VolMax
                ws.Range("P4").Value = VolMaxTicker
                ws.Range("Q4").NumberFormat = "0.0000E+00"
            End If
                
        Next j
        
'CHALLENGE find values for this worksheet and move onto the next
Next ws
        
End Sub
