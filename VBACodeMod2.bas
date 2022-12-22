Attribute VB_Name = "Module1"
Sub StockValues()

'---------------------------
'Loop Through All Sheets
'---------------------------

For Each ws In Worksheets

    '----------------------------
    'Create variable to Hold File Name, Last Row
    Dim WorksheetName As String
    
    'Determine the LastRow
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Grabbed the WorksheetName
    WorksheetName = ws.Name
    'Msgbox WorksheetName
    
    'Add a Column for Ticker
    ws.Range("I1").EntireColumn.Insert
    
    'Add the word "Ticker" to the "I" Column Header
    ws.Cells(1, 9).Value = "Ticker"
    
    'Add a Column for Yearly Change
    ws.Range("J1").EntireColumn.Insert
    
    'Add the words "Yearly Change" to the "J" Column Header
    ws.Cells(1, 10).Value = "Yearly Change"
    
    'Add a Column for Ticker
    ws.Range("K1").EntireColumn.Insert
    
    'Add the words "Percent Change" to the "K" Column Header
    ws.Cells(1, 11).Value = "Percent Change"
    
    'Add the Column for Total Stock Volume
    ws.Range("L1").EntireColumn.Insert
    
    'Add the words "Total Stock Volume" to the "K" Column Header
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    '---------------------------------
    'Loop Through Ticker, Open, Close, and Total Stock Volume
    '---------------------------------
    
        'Set an initial variable for holding the ticker
        Dim Ticker2 As String
        
        'Set an initial variable for holding the open value
        Dim beginprice As Double
        
        'Set an initial variable for holding the closing value
        Dim endprice As Double
        
        'Set an initial variable for holding the total stock volume
        Dim Volume_Total As LongLong
        
        'Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        
        'Loop through all rows
         For i = 2 To Lastrow
        
        
            'Check if we are still within the same ticker, if it is not...
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
                'Set ticker name
                Ticker2 = ws.Cells(i, 1).Value
            
                'Add to the total stock volume
                Volume_Total = Volume_Total + Cells(i, 7).Value
                
                'Set open value
                beginprice = ws.Cells(i, 3).Value
             
                'Set close value
                endprice = ws.Cells(i, 6).Value
             
                'Formulate the yearly and percent change
                Change = endprice - beginprice
             
                Perchange = endprice - beginprice / beginprice * 100
                
                'Percent formatting
                ws.Cells(i, 11).Value = Format(Perchange, "Percent")
                
                'Else
                
                ws.Cells(i, 11).Value = Format(0, "Percent")
             
            
                'Print Ticker in Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker2
                
                'Print Yearly Change in Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Change
                
                'Print Percent Change in Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Perchange
            
                'Print total stock volume to the summary table
                ws.Range("L" & Summary_Table_Row).Value = Volume_Total
                
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    End If
            
                'Reset the data
                Volume_Total = 0
                beginprice = 0
                endprice = 0
            
                
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'If the cell immediately following a row is the same ticker...
                Else
                
                
                'Add to the total stock volume
                Volume_Total = Volume_Total + Cells(i, 7).Value
                 
                End If
                
            Next i
            
    'Add place for greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Get the last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Start variables and set values of variables to the first row
    greatest_percent_increase = ws.Cells(2, 11).Value
    greatest_percent_increase_ticker = ws.Cells(2, 9).Value
    greatest_percent_decrease = ws.Cells(2, 11).Value
    greatest_percent_decrease_ticker = ws.Cells(2, 9).Value
    greatest_stock_volume = ws.Cells(2, 12).Value
    greatest_stock_volume_ticker = ws.Cells(2, 9).Value
    
    
    'loop through the list
    For i = 2 To lastRowState
    
        'Find the ticker with the greatest percent increase
        If ws.Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = ws.Cells(i, 11).Value
            greatest_percent_increase_ticker = ws.Cells(i, 9).Value
        End If
        
        'Find the ticker with the greatest percent decrease
        If ws.Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = ws.Cells(i, 11).Value
            greatest_percent_decrease_ticker = ws.Cells(i, 9).Value
        End If
        
        'Find the ticker with the greatest stock volume
        If ws.Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = ws.Cells(i, 12).Value
            greatest_stock_volume_ticker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    'Add the values for greatest percent increase, decrease, and stock volume to each worksheet
    ws.Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    ws.Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    ws.Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    ws.Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    ws.Range("P4").Value = greatest_stock_volume_ticker
    ws.Range("Q4").Value = greatest_stock_volume
    
        Next ws
        
            
        
        
        
            
            
            
            
            
        
        
        
        
    
    
    
    
    
End Sub
