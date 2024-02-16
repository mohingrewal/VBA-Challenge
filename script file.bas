Attribute VB_Name = "Module1"
Sub Stockanalysis()

'Loop through entire workbook
For Each ws In Worksheets

    'Let us set an initial variable for For loop
    Dim j As Long
    'Set variable to keep track of location for each ticker in the summary table
    Dim ticker_value As String
    
     'Define loop for hard part
    Dim i As Double
    'Set variable to keep track of location for each ticker in the summary table
    Dim the_ticker As String
    
    'Set variable to keep track of total volume
    Dim allvolume As Double
        allvolume = 0
    
    'Set variable to keep track of ticker_row for the summary
    Dim smytick As Double
        smytick = 2
    
    'Set variables for yearly change
    Dim opening_price As Double
    Dim closing_price As Double
    Dim Change As Double
    
    'Calculate last row for j
    Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'create a couple more strings to save (j,1) and (j +1,1) as they are referenced so many times
    Dim one, two As String
    
    'Start loop here
    For j = 2 To lastrow
        
    'let's save (j,1) and (j+1,1) at the starting of our loop
    current_1 = ws.Cells(j, 1).Value
    next_1 = ws.Cells(j + 1, 1).Value
    prior_1 = ws.Cells(j - 1, 1).Value
    
    'Calculate total volume
    allvolume = allvolume + ws.Cells(j, 7).Value
            
            'In this case we can combine our logic into a single if/else
            If current_1 <> next_1 Then
                closing_price = ws.Cells(j, 6).Value
                
                'If the loop gets to this point, the opening price will be received
                Change = closing_price - opening_price
                
                'Create a nested If statement to calculate the % change with the opening price and closing price
                ws.Cells(smytick, 10).Value = Change
                    If opening_price <> 0 Then
                'The divisor cannot be 0 but the numerator can
                        percentage = Change / opening_price
                        ws.Range("K" & smytick).Value = percentage
                    End If
                
                'Place total volume and ticker to stock summary
                ws.Range("I" & smytick).Value = current_1
                ws.Range("L" & smytick).Value = allvolume
                'Add a row for next ticker
                smytick = smytick + 1
                'Reset volume to 0 each time a new ticker is found'
                allvolume = 0
                
             ElseIf previous_1 <> current_1 Then
                opening_price = ws.Cells(j, 3).Value
             End If
        Next j
        
        'Move formatting outside of the loop
        ws.Range("K:K").EntireColumn.NumberFormat = "0.00%"
        ws.Range("L:L").EntireColumn.NumberFormat = "0"
    
   
    
    'Find the greatest increase, decrease, and volume
    Greatest_Increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 16).Value = Greatest_Increase
    Greatest_Decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 16).Value = Greatest_Decrease
    Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 16).Value = Greatest_Volume
    
    'Begin loop
        For i = 2 To lastrow
        Condiforma = ws.Cells(i, 11).Value
        the_ticker = ws.Cells(i, 9).Value
        
            If ws.Cells(i, 11).Value = Greatest_Increase Then
                ws.Cells(2, 15).Value = the_ticker
                
            ElseIf ws.Cells(i, 11).Value = Greatest_Decrease Then
                ws.Cells(3, 15).Value = the_ticker
                
            ElseIf ws.Cells(i, 12).Value = Greatest_Volume Then
                ws.Cells(4, 15).Value = the_ticker
                
            End If
            
            'Conditionally format the percent change
            If Condiforma < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 5
            ElseIf Condiforma > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 7
            End If
    Next i
    'Change P2 and P3 to percentage format
    ws.Range("P2:P3").NumberFormat = "0.00%"

    'Format worksheets
        'Make column titles
        ws.Range("I1").Value = "Ticker_Value"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'Make row titles
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        ws.Range("A:P").EntireColumn.AutoFit
Next ws
End Sub
