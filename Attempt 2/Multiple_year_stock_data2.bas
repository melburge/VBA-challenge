Attribute VB_Name = "Module1"
Sub StockData()

 ' Create Variables to store the values
 Dim tickerSymbolName As String
 Dim stockVolume As Long
 stockVolumeTotal = 0
 
 Dim summaryTableRow As Integer
 summaryTableRow = 2
 
 Dim openPrice As Double
 Dim column As Integer
 Dim closePrice As Double
 Dim yearlyChange As Double
 
 Dim lastrowK As Double
 Dim lastrowL As Long
 
 Dim percentIncrease As Double
 Dim tickerPercentIncrease As Integer
 Dim percentDecrease As Double
 Dim tickerPercentDecrease As Integer
 Dim grtTotalVolume As Double
 Dim tickergrtTotalVolume As Integer
 grtTotalVolume = 0
 
   
 Dim ws As Worksheet
    
    For Each ws In Worksheets
      ws.Activate
    
        summaryTableRow = 2
        ' Create Columns to store the data
        Range("I1").Value = "Ticker Symbol"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"
    
        ' Bonus columns and cells
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' Autofit the columns
        Columns("I:L").EntireColumn.AutoFit
    
        'Change column K to percentage
        Columns("K:K").EntireColumn.NumberFormat = "0.00%"
            
        ' Declare the Variables with information
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        lastrowK = ws.Cells(Rows.Count, "K").End(xlUp).Row
        lastrowL = ws.Cells(Rows.Count, "L").End(xlUp).Row
        openPrice = ws.Cells(2, "C").Value
        
        ' loop through Data Set to obtain values
        For i = 2 To lastrow
                  
            ' Add stock volumns
            stockVolumeTotal = stockVolumeTotal + ws.Cells(i, "G").Value
                                
            ' Check if ticker symbol name is the same
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
                                
                ' Ticker Symbol name
                tickerSymbolName = ws.Cells(i, "A").Value
                
                ' Obtain the open and close price and enter difference plus %
                closePrice = ws.Cells(i, "F").Value
                yearlyChange = closePrice - openPrice
                percentageChange = yearlyChange / openPrice
                           
                ' Print Ticker Symbol, Yearly Change, % Change and Stock Volume Total in summary table
                ws.Range("I" & summaryTableRow).Value = ws.Cells(i, "A").Value
   
                ws.Range("J" & summaryTableRow).Value = yearlyChange
             
                ws.Range("K" & summaryTableRow).Value = percentageChange
            
                ws.Range("L" & summaryTableRow).Value = stockVolumeTotal
                
            
                ' Change Cell colour
                If yearlyChange > 0 Then
                ws.Cells(summaryTableRow, "J").Interior.ColorIndex = 4
                
                ElseIf yearlyChange < 0 Then
                ws.Cells(summaryTableRow, "J").Interior.ColorIndex = 3
                    
                Else
                ws.Cells(summaryTableRow, "J").Interior.ColorIndex = 2
                       
                End If
                
                ' Add a new line to the summary table
                summaryTableRow = summaryTableRow + 1
                
                ' Reset Stock Volume
                stockVolumeTotal = 0
                openPrice = ws.Cells(i + 1, "C").Value
                                                     
                ' Summary Table 2 Greatest Increase, Decrease andTotal volume
                Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrowK)) * 100
                Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrowK)) * 100
                Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrowL))
                
                'return one less because header row is not a factor'
                increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrowK)), Range("K2:K" & lastrowK), 0)
                decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrowK)), Range("K2:K" & lastrowK), 0)
                volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrowL)), Range("L2:L" & lastrowL), 0)
                
                Range("P2") = Cells(increase_number + 1, "A")
                Range("P3") = Cells(decrease_number + 1, "A")
                Range("P4") = Cells(volume_number + 1, "A")
                
                Columns("O:Q").EntireColumn.AutoFit
                
            End If
        Next i
        
    Next ws
    
    MsgBox "Complete"
    
End Sub





