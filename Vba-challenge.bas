Attribute VB_Name = "Module1"
Sub Analyzer()
    
    Dim ws As Worksheet
    
    'Loop through all worksheets
    For Each ws In Worksheets

        'define variables
        Dim Ticker As String
    
        Dim Stock_Total As LongLong
        Stock_Total = 0
    
        Dim Percent_Change As Double
        Percent_Change = 0
    
        Dim Price_Change As Double
        Price_Change = 0
    
        Dim Sum_Table_Row As Integer
        Sum_Table_Row = 2
    
        Dim Open_Price As Double
        Open_Price = 0
    
        Dim Close_Price As Double
        Close_Price = 0
        
        Dim LR As Long
        LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'create new variables for bonus
        Dim Great_percent_inc As Double
        Great_percent_inc = 0
        
        Dim Great_percent_dec As Double
        Great_percent_dec = 0
        
        Dim Greatest_Total As LongLong
        Greatest_Total = 0
        
        
        'Set initial open_price
        Open_Price = ws.Cells(2, 3)
    
        'create table headers
        ws.Range("I1").Value = ("Ticker")
        ws.Range("J1").Value = ("Yearly Change")
        ws.Range("K1").Value = ("Percent Change")
        ws.Range("L1").Value = ("Total Stock Volume")
            
        'create bonus headers
        ws.Range("P1").Value = ("Ticker")
        ws.Range("Q1").Value = ("Value")
        ws.Range("O2").Value = ("Greatest % Increase")
        ws.Range("O3").Value = ("Greatest % Decrease")
        ws.Range("O4").Value = ("Greatest Total Volume")
        
        'create for loop
        For i = 2 To LR
        
            'Check if still in same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'set ticker name
                Ticker = ws.Cells(i, 1).Value
            
                'Set close price
                Close_Price = ws.Cells(i, 6).Value
        
                'Calculate price change
                Price_Change = Close_Price - Open_Price
            
                'Calculate percentage change
                If Price_Change <> 0 Then
                    Percent_Change = (Price_Change / Open_Price)
                Else
                    Percent_Change = 0
                End If
        
                'add to stock total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
                'Print ticker name in table
                ws.Range("i" & Sum_Table_Row).Value = Ticker
            
                'Print Yearly Change
                ws.Range("j" & Sum_Table_Row).Value = Price_Change
                ws.Range("j" & Sum_Table_Row).NumberFormat = "0.00"
            
                'Print Percent Change
                ws.Range("K" & Sum_Table_Row).Value = Percent_Change
                ws.Range("K" & Sum_Table_Row).NumberFormat = "0.00%"
            
                'Print stock total
                ws.Range("l" & Sum_Table_Row).Value = Stock_Total
            
                'Set Color conditionals for yearly change
                If ws.Range("J" & Sum_Table_Row).Value >= 0 Then
                    ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 3
                End If
            
                'set new open price
                j = i + 1
                Open_Price = ws.Cells(j, 3).Value
            
                'add one to summary table
                Sum_Table_Row = Sum_Table_Row + 1
            
                'Reset stock total
                Stock_Total = 0
            
            Else
             
                'add to stock total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
        
            End If
       
        Next i
        
        'Create new ticker variable for Bonus
        Dim Tick1 As String
        Tick1 = 0
        
        Dim Tick2 As String
        Tick2 = 0
        
        Dim Tick3 As String
        Tick3 = 0
        
        'Create for loop for Bonus
        For i = 2 To LR
        
            If ws.Cells(i, 11).Value > Great_percent_inc Then
                
                Great_percent_inc = ws.Cells(i, 11).Value
                Tick1 = ws.Cells(i, 9).Value
                
            End If
            
            If ws.Cells(i, 11).Value < Great_percent_dec Then
            
                Great_percent_dec = ws.Cells(i, 11).Value
                Tick2 = ws.Cells(i, 9).Value
            
            End If
            
            If ws.Cells(i, 12).Value > Greatest_Total Then
            
                Greatest_Total = ws.Cells(i, 12).Value
                Tick3 = ws.Cells(i, 9).Value
            
            End If
                
        Next i
        
        'Print Bonus tickers
        ws.Range("P2").Value = Tick1
        ws.Range("P3").Value = Tick2
        ws.Range("P4").Value = Tick3
        
        'print bonus values
        ws.Range("Q2").Value = Great_percent_inc
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("Q3").Value = Great_percent_dec
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("Q4").Value = Greatest_Total
        
                
    
    Next ws
    
End Sub
