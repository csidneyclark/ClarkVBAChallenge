<<<<<<< HEAD
# ClarkVBAChallenge

Sub StockMarket()
  
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

        ' Headers for Summary Table
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        'Name variables
       
        Dim ticker_type As String
        Dim ticker_total As Double
        Dim LastRow As Long
            LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Dim Open_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim summary_table_row As Double
            summary_table_row = 2
            
        
        ' Setting ticker_total at 0
        
        ticker_total = 0
        
        'Defining opening price
        Open_Price = Cells(2, 3).Value
       
       'Create for loop
       
        
        For i = 2 To LastRow
        
            'If the cell directly below is a different ticker, then determine original cell's ticker type
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker_type = Cells(i, 1).Value
                
                'Print the ticker type
                Cells(summary_table_row, 9).Value = ticker_type
                
                ' Defining Closing price
                Closing_Price = Cells(i, 6).Value
                
                ' Calculating yearly change
                
                Yearly_Change = Closing_Price - Open_Price
                
                'Printing yearly change
                
                Cells(summary_table_row, 10).Value = Yearly_Change
                
                ' Calculating percent change
                
                        If (Open_Price = 0 And Closing_Price = 0) Then
                            Percent_Change = 0
                        ElseIf (Open_Price = 0 And Closing_Price <> 0) Then
                            Percent_Change = 1
                        Else
                            Percent_Change = Yearly_Change / Open_Price
                    
                'Printing percent change
                
                Cells(summary_table_row, 11).Value = Percent_Change
                Cells(summary_table_row, 11).NumberFormat = "0.00%"
                        
                        End If
                
                'ticker_total is the running total plus the value of total in current cell
                
                ticker_total = ticker_total + Cells(i, 7).Value
                
                Cells(summary_table_row, 12).Value = ticker_total
                
                ' Move down a row in summary table
                
                summary_table_row = summary_table_row + 1
                
                ' reset the Open Price
                Open_Price = Cells(i + 1, 3)
                
                ' reset the ticker total
                ticker_total = 0
                
            
            
            Else
                'if cells are the same ticker, ticker_total is the running total plus the value of total in current cell
                ticker_total = ticker_total + Cells(i, 7).Value
            
            End If
            
        Next i
        
        ' Last row for yearly change
        
        lastrow_yearlychange = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Cell Colors for yearly change
        
        For i = 2 To lastrow_yearlychange
            If Cells(i, 10).Value > 0 Then
                    
                Cells(i, 10).Interior.ColorIndex = 4
                    
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i
        
                
    Next WS
        
=======
# ClarkVBAChallenge

Sub StockMarket()
  
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

        ' Headers for Summary Table
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        'Name variables
       
        Dim ticker_type As String
        Dim ticker_total As Double
        Dim LastRow As Long
            LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Dim Open_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim summary_table_row As Double
            summary_table_row = 2
            
        
        ' Setting ticker_total at 0
        
        ticker_total = 0
        
        'Defining opening price
        Open_Price = Cells(2, 3).Value
       
       'Create for loop
       
        
        For i = 2 To LastRow
        
            'If the cell directly below is a different ticker, then determine original cell's ticker type
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ticker_type = Cells(i, 1).Value
                
                'Print the ticker type
                Cells(summary_table_row, 9).Value = ticker_type
                
                ' Defining Closing price
                Closing_Price = Cells(i, 6).Value
                
                ' Calculating yearly change
                
                Yearly_Change = Closing_Price - Open_Price
                
                'Printing yearly change
                
                Cells(summary_table_row, 10).Value = Yearly_Change
                
                ' Calculating percent change
                
                        If (Open_Price = 0 And Closing_Price = 0) Then
                            Percent_Change = 0
                        ElseIf (Open_Price = 0 And Closing_Price <> 0) Then
                            Percent_Change = 1
                        Else
                            Percent_Change = Yearly_Change / Open_Price
                    
                'Printing percent change
                
                Cells(summary_table_row, 11).Value = Percent_Change
                Cells(summary_table_row, 11).NumberFormat = "0.00%"
                        
                        End If
                
                'ticker_total is the running total plus the value of total in current cell
                
                ticker_total = ticker_total + Cells(i, 7).Value
                
                Cells(summary_table_row, 12).Value = ticker_total
                
                ' Move down a row in summary table
                
                summary_table_row = summary_table_row + 1
                
                ' reset the Open Price
                Open_Price = Cells(i + 1, 3)
                
                ' reset the ticker total
                ticker_total = 0
                
            
            
            Else
                'if cells are the same ticker, ticker_total is the running total plus the value of total in current cell
                ticker_total = ticker_total + Cells(i, 7).Value
            
            End If
            
        Next i
        
        ' Last row for yearly change
        
        lastrow_yearlychange = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Cell Colors for yearly change
        
        For i = 2 To lastrow_yearlychange
            If Cells(i, 10).Value > 0 Then
                    
                Cells(i, 10).Interior.ColorIndex = 4
                    
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i
        
                
    Next WS
        
>>>>>>> fbc60e35c520e45812a70a56ab0410d8f677469d
End Sub