Attribute VB_Name = "Module1"
'This Subroutine runs on every single Worksheet of an active workbook and populates data from Column I to Column Q
'It calculates Yearly Change, Percent Change, Total Stock Volume,Greatest % Increase, Greatest % Decrease and Greatest Total Volume on a given Stocks data

Sub Analyse_Stocks_Data()
    
    Dim ws As Worksheet
    
    'Loop through all the worksheets
    For Each ws In ThisWorkbook.Worksheets
        'Make current worksheet active so that changes are done on current worksheet
        Worksheets(ws.Name).Activate
        
        'Declaring and initializing variables
        Dim Ticker As String
        Dim Yearly_Change As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Dim Max_Ticker_Name As String
        Dim Min_Ticker_Name As String
        Dim Max_Percent As Double
        Dim Min_Percent As Double
        Dim Max_Ticker_Volume As String
        Dim Max_Volume As Double
        
        Total_Stock_Volume = 0
        Max_Percent = 0
        Min_Percent = 0
        Max_Volume = 0
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Ticker = Cells(2, 1).Value
        Open_Price = Cells(2, 3).Value
        
        'Search and store last Row number
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 2 To lastrow
        
            'if current ticker is not equal to next ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'perform calculations for current ticker
                Ticker = Cells(i, 1).Value
               
                Close_Price = Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                Percent_Change = Yearly_Change / Open_Price
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                 
                'Table headings
                Cells(1, 9).Value = "Ticker"
                Cells(1, 10).Value = "Yearly Change"
                Cells(1, 11).Value = "Percent Change"
                Cells(1, 12).Value = "Total Stock Volume"
                Range("I1:L1").Font.Bold = True
                
                Cells(2, 15).Value = "Greatest % Increase"
                Cells(3, 15).Value = "Greatest % Decrease"
                Cells(4, 15).Value = "Greatest Total Volume"
                Range("O2:O4").Font.Bold = True
                
                Cells(1, 16).Value = "Ticker"
                Cells(1, 17).Value = "Value"
                Range("P1:Q1").Font.Bold = True
                
                'Enter calculated values in the cells
                Range("I" & Summary_Table_Row).Value = Ticker
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                Range("J" & Summary_Table_Row).NumberFormat = "[$$-en-CA] #,##0.00"
                Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change, 2)
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                Range("P2").Value = Max_Ticker_Name
                Range("P3").Value = Min_Ticker_Name
                Range("P4").Value = Max_Ticker_Volume
                Range("Q2").Value = FormatPercent(Max_Percent, 2)
                Range("Q3").Value = FormatPercent(Min_Percent, 2)
                Range("Q4").Value = Max_Volume
                
                'Conditional formatting for negative / positive values
                If (Yearly_Change > 0) Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ElseIf (Yearly_Change < 0) Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                End If
                
                
                Summary_Table_Row = Summary_Table_Row + 1
                Open_Price = Cells(i + 1, 3).Value
                
                'Update greatest % Increase and greatest % Decrease Value with ticker name
                If Percent_Change > Max_Percent Then
                    Max_Percent = Percent_Change
                    Max_Ticker_Name = Ticker
                    
                ElseIf Percent_Change < Min_Percent Then
                    Min_Percent = Percent_Change
                    Min_Ticker_Name = Ticker
                    
                End If
                
                'Update Greatest Total Volume value with ticker name
                If (Total_Stock_Volume > Max_Volume) Then
                    Max_Volume = Total_Stock_Volume
                    Max_Ticker_Volume = Ticker
                    
                End If
                
                'Reset value for next ticker
                Total_Stock_Volume = 0
            Else
            
                'Updating total_stock_volume for current ticker
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            End If
  
            Next i
 
    Next ws

End Sub




