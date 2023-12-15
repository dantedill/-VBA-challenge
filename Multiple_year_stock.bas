Attribute VB_Name = "Module1"
Sub StockInfo():

        'assign Variables
        Dim total As Double 'Holds Total Stock Volume
        Dim yearlyChange As Double 'Holds yearly change
        Dim percentChange As Double 'Percent change of each stock
        Dim summaryTableRow As Long  'Holds row of Summary Table
        Dim newStockRow As Long 'Holds value of next stock list
        Dim row As Long 'Loop through rows
        Dim rowCount As Long 'Hold value of row numbers
        
        
        ' loop all worksheets
        For Each ws In Worksheets
        
        ' Set up start values
        total = 0 'Total Stock Volume start at 0
        yearlyChange = 0 'Yearly Change start at 0
        summaryTableRow = 0 'Summary Tablerow start at 0
        newStockRow = 2 'sheet info start on row 2
        
        rowCount = Cells(Rows.Count, "A").End(xlUp).row ' Holds last row value
        
        
            'Set the tile cols
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ws.Range("P1").Value = "Ticker"
            ws.Range("O1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest 5 Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
                'Start loop
                For row = 2 To rowCount 'Loop to end
                    
                    ' checking to see if column matches
                    If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                    
                        total = total + ws.Cells(row, 7).Value 'Pull value in Colmn G
                            'checking to see if toatal vol = 0
                            
                            If total = 0 Then
                             'print results in Columns I, J,K'and L
                             ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value 'print stock name in cols I
                             ws.Range("J" & 2 + summaryTableRow).Value = 0 'print 0 in cols J-Yearly Change
                             ws.Range("K" & 2 + summaryTableRow).Value = 0 'print 0 in cols K-0% Change
                             ws.Range("L" & 2 + summaryTableRow).Value = 0 'print 0 in cols L-Total stock volume
                            Else
                            
                                'Looking for zero starting value
                                If ws.Cells(newStockRow, 3).Value = 0 Then
                                    For findValue = newStockRow To row
                                     'check to see if the next (or next) value does not equal 0
                                    If ws.Cells(findValue, 3).Value <> 0 Then
                                     newStockRow = findValue
                                    ' once we ave a non-zero value, breakou od the loop
                                    Exit For
                                    End If
                                  Next findValue
                               End If
                               
                                    'Calculate the year change last close and first open
                                    yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(newStockRow, 3).Value)
                                    'calculate the percent change (yearly change / first open
                                    percentChange = yearlyChange / ws.Cells(newStockRow, 3).Value
                                    
                            'print results in Columns I, J,K'and L
                             ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value 'print stock name
                             ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange 'print 0 in cols J-Yearly Change
                             ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00" 'formats J-Yearly Change
                             ws.Range("K" & 2 + summaryTableRow).Value = percentChange 'formats K-0% Change
                              ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00" 'formats J-Yearly Change
                             ws.Range("L" & 2 + summaryTableRow).Value = total 'print 0 in cols L-Total stock volume
                             ws.Range("J" & 2 + summaryTableRow).NumberFormat = "#,###" 'formats J-Yearly Change
                             
                                    ' formatting cells with color year change
                                    If yearlyChange > 0 Then
                                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' green if pos
                                        ElseIf yearlyChange < 0 Then
                                            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' red if neg
                                        Else
                                            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0  'white if zero
                                    End If
                            End If
                            total = 0 ' reset total(Total Vloume) to 0
                            ' Move summary Table to next row
                            yearlyChange = 0
                            summaryTableRow = summaryTableRow + 1
                            
                            
                    Else  'if ticker is the same
                     total = total + Cells(row, 7).Value 'Pull value in Colmn G
                
                   End If
                
                
                Next row
                    'After all rows look for Max and Min in Q2, Q3, Q4
                    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100  'greatest increase
                    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100  'greatest decrease
                    ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100   'Greatest max volume
                    ws.Range("Q4").NumberFormat = "#.###" 'Formats
                    
                    increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
                    ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
                    
                      decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
                    ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
                    
                    greatVolumeNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
                    ws.Range("P4").Value = ws.Cells(decreaseNumber + 1, 9)
                    
                    
                    'auto fit Colms
            ws.Columns("A:Q").AutoFit
    
Next ws

End Sub
