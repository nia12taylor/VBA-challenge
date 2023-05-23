Attribute VB_Name = "Module1"
Sub yearlyStockAnalysis()
  Dim ws As Worksheet

'Print two new tables on each sheet to summarize that sheet's stock data

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        'Print the headers for table one
        
            'Print the Ticker column label
            ws.Range("I1").Value = "Ticker"
            
            'Print the Yearly Change column label
            ws.Range("J1").Value = "Yearly Change"
            
            'Print the Percentage Change column label
            ws.Range("K1").Value = "Percentage Change"
            
            'Print the Total Stock Volume column label
            ws.Range("L1").Value = "Total Stock Volume"
            
        'Print the headers for table two
            
            'Print the header for the Ticker column
            ws.Range("P1").Value = "Ticker"
            
            'Print the header for the Value column
            ws.Range("Q1").Value = "Value"
            
            'Print the header for the Greatest % Increase row
            ws.Range("O2").Value = "Greatest % Increase"
            
            'Print the header for the Greatest % Decrease row
            ws.Range("O3").Value = "Greatest % Decrease"
            
            'Print the header for the Greatest Total Volume row
            ws.Range("O4").Value = "Greatest Total Volume"
            
            'Auto adjust the width of the row headers
            ws.Columns("O").AutoFit
        
        'Populate table one
                  
            'Count the number of rows in the dataset
            Dim rowcount As Long
            rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'Keep track of the row in table one
            Dim tickercount As Long
            tickercount = 2
            
            'Store the current <ticker> from the dataset
            Dim currentticker As String
            currentticker = ws.Cells(2, 1).Value
            
            'Store the <vol> sum for the current <ticker>
            'This variable needs to have data type LongLong or the value will overflow; use of LongLong requires Excel 64-bit
            Dim volsum As LongLong
            volsum = 0
            
            'Store the first <open> of the current <ticker>
            Dim firstopen As Double
            firstopen = ws.Cells(2, 3).Value
            
            'Store the last <close> of the current <ticker>
            Dim lastclose As Double
            lastclose = 0
            
            'Store the Greatest % Increase
            Dim greatestincrease As Double
            greatestincrease = 0
            
            'Store the Ticker with the Greatest % Increase
            Dim tickerincrease As String
            tickerincrease = 0
            
            'Store the Greatest % Decrease
            Dim greatestdecrease As Double
            greatestdecrease = 0
            
            'Store the Ticker with the Greatest % Decrease
            Dim tickerdecrease As String
            tickerdecrease = 0
            
            'Store the Greatest Total Volume
            Dim greatestvol As LongLong
            greatestvol = 0
            
            'Store the ticker with the Greatest Total Volume
            Dim tickervol As String
            tickervol = ws.Cells(2, 9).Value
        
                'Loop through the rows in the dataset
                For a = 2 To rowcount
            
                    'Add the <vol> of the current row to <vol> sum
                    volsum = volsum + ws.Cells(a, 7).Value
                
                    'If the current <ticker> is not the same as the next <ticker>
                    If currentticker <> ws.Cells(a + 1, 1).Value Then
        
                        'Print the current <ticker> in the Ticker column of table one
                        ws.Cells(tickercount, 9).Value = currentticker
                        
                        'Store the last <close> of the current <ticker>
                        lastclose = ws.Cells(a, 6).Value
                                              
                        'Calculate the Yearly Change using this formla:  Yearly Change = (last <close> - first <open>)
                        'Format the result by giving it two decimal places
                        'Print the result in the Yearly Change column of table one
                        ws.Cells(tickercount, 10).Value = FormatNumber(lastclose - firstopen, 2)
                        'Format the result with Excel's "Number" format
                        ws.Cells(tickercount, 10).NumberFormat = "0.00"
                        
                        'Update the current <ticker>
                        currentticker = ws.Cells(a + 1, 1).Value
                        
                            'If the Yearly Change is positive
                            If ws.Cells(tickercount, 10).Value > 0 Then
                    
                                'Fill the cell with a green color
                                ws.Cells(tickercount, 10).Interior.ColorIndex = 4
                        
                            'Else if the Yearly Change is negative
                            ElseIf ws.Cells(tickercount, 10).Value < 0 Then
                            
                                'Fill the cell with a red color
                                ws.Cells(tickercount, 10).Interior.ColorIndex = 3
                    
                            End If
                        
                        'Calculate the Percentage Change using this formula: Percentage Change = ((last <close>/first <open>)-1)
                        'Format the result with Excel's "Percentage" format
                        'Print the result in the Percentage Change column of table one
                        ws.Cells(tickercount, 11).Value = FormatPercent((lastclose / firstopen) - 1)
                        
                            'If the Percentage Change is bigger than the current Greatest % Increase
                            If ws.Cells(tickercount, 11).Value > greatestincrease Then
                            
                                'Store that Percentage Change as the new Greatest % Increase
                                greatestincrease = ws.Cells(tickercount, 11).Value
                                
                                'Store the Ticker associated with that Percentage Change
                                tickerincrease = ws.Cells(tickercount, 9).Value
                            
                            'Else if the next Percentage Change is smaller than the current Greatest % Decrease
                            ElseIf ws.Cells(tickercount, 11).Value < greatestdecrease Then
                            
                                'Store that Percentage Change as the new Greatest % Decrease
                                greatestdecrease = ws.Cells(tickercount, 11).Value
                                
                                'Store the Ticker associated with that Percentage Change
                                tickerdecrease = ws.Cells(tickercount, 9).Value
                
                            End If
                        
                        'Print the <vol> sum in the Total Stock Volume column of table one
                        ws.Cells(tickercount, 12).Value = volsum
                        
                            'If the <vol> sum that was just printed is bigger than the current Greatest Total Volume
                            If ws.Cells(tickercount, 12).Value > greatestvol Then
                            
                                'Store the value of that <vol> sum as the new Greatest Total Volume
                                greatestvol = ws.Cells(tickercount, 12).Value
                                
                                'Store the ticker associated with that <vol> sum
                                tickervol = ws.Cells(tickercount, 9).Value
                
                            End If
                            
                        'Store the first <open> of the new <ticker>
                        firstopen = ws.Cells(a + 1, 3).Value
                        
                        'Reset the <vol> sum
                        volsum = 0
                        
                        'Shift down by one row in table one
                        tickercount = tickercount + 1
                
                    End If
                    
                Next a
                
            'Auto adjust the width of the columns in table one
            ws.Columns("J:L").AutoFit
            
        'Populate table two
            
            'Print the ticker with the Greatest % Increase
            ws.Range("P2").Value = tickerincrease
            
            'Print the value of the Greatest % Increase
            ws.Range("Q2").Value = FormatPercent(greatestincrease)
            
            'Print the ticker with the Greatest % Decrease
            ws.Range("P3").Value = tickerdecrease
                
            'Print the value of the Greatest % Decrease
            ws.Range("Q3").Value = FormatPercent(greatestdecrease)
            
            'Print the ticker with the Greatest Total Volume
            ws.Range("P4").Value = tickervol
            
            'Print the value of the Greatest Total Volume
            ws.Range("Q4").Value = greatestvol
    
    Next ws
    
    MsgBox ("Populate complete.")

End Sub
Sub Reset():

'This subroutine is for testing purposes only. It resets all sheets in the Workbook back to their pre-populate state.

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

    'Reset the worksheet to its original state
    
        'Clear the cells
        ws.Columns("I:L").Clear
        ws.Columns("O:Q").Clear
        
        'Reset the widths of the columns
        ws.Columns("J:L").ColumnWidth = 8.43
        ws.Columns("O").ColumnWidth = 8.43
        
        'Remove the number formats
        ws.Range("J2:J91").ClearFormats
    
    Next ws
    
    MsgBox ("Reset complete.")

End Sub
