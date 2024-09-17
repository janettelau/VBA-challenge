Attribute VB_Name = "Module1"
Sub AlphabeticalTesting()

    ' Declaring variables
    
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim ticker As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim total_volume As Double
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim start As Long
    
    ' Create a variable to hold the counter
    Dim i As Long
    
    ' Variables for tracking the greatest increase, decrease, and total volume
    Dim max_increase As Double
    Dim max_increase_ticker As String
    Dim max_decrease As Double
    Dim max_decrease_ticker As String
    Dim max_volume As Double
    Dim max_volume_ticker As String
    
    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' Keep track of the location for each stock in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Initialize tracking variables
        max_increase = 0
        max_decrease = 0
        max_volume = 0
    
        ' Write headers for output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ' Counts the number of rows
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set the First Row
        start = 2
    
        For i = 2 To lastrow
            
            ' Check if we are still within the same stock, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set the Ticker Name
                ticker = ws.Cells(i, 1).Value
                
                ' Set the Opening Price
                opening_price = ws.Cells(start, 3).Value
                
                ' Set Closing Price
                closing_price = ws.Cells(i, 6).Value
                
                ' Calculate the Quarterly Change
                quarterly_change = closing_price - opening_price
                
                ' Calculate the Total Volume
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                ' Calculate the Percentage Change
                percent_change = (quarterly_change) / opening_price
    
                 ' Print the Ticker Name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker
    
                ' Print the Quarterly Change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = quarterly_change
                ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
                
                ' Check if the Quarterly Change is Positive
                If quarterly_change > 0 Then

                    ' Color the Positive Change green
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                ' Check if the Quarterly Change is Negative
                ElseIf quarterly_change < 0 Then
                
                    ' Color the Negative Change red
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                ' Print the Percentage Change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ' Print the Total Volume in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = total_volume
                
                ' Check for greatest percentage increase
                If ws.Range("K" & Summary_Table_Row).Value > max_increase Then
                    max_increase = percent_change
                    max_increase_ticker = ticker
                End If
                
                ' Check for greatest percentage decrease
                If ws.Range("K" & Summary_Table_Row).Value < max_decrease Then
                    max_decrease = percent_change
                    max_decrease_ticker = ticker
                End If
                
                ' Check for greatest total volume
                If total_volume > max_volume Then
                    max_volume = total_volume
                    max_volume_ticker = ticker
                End If
                
                ' Reset the Total Volume for the next Ticker
                total_volume = 0
                
                ' Update the Starting Row
                start = i + 1
                
                ' Update the Opening Price for the next Ticker
                opening_price = ws.Cells(start, 3).Value
    
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
    
            ' If the cell immediately following a row is the same stock
            Else
    
                ' Add to the Total Volume
                total_volume = total_volume + ws.Cells(i, 7).Value
    
            End If
    
        Next i
        
        ' Print the results for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("Q2").Value = max_increase
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("Q3").Value = max_decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
        
        ' Adjust Column Widths using AutoFit
        ws.Range("J1:L1").Columns.AutoFit
        ws.Columns("O").AutoFit
        
    Next ws

End Sub

