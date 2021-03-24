Attribute VB_Name = "Module1"
Sub Stocktickers()
' setting variables
Dim ticker As String
Dim open_price, closing_price As Double
Dim yearly_change, percent_change As Double
Dim Total_stock_volume As LongLong

' Loop through all sheets of the whorksheets
For Each ws In ThisWorkbook.Worksheets

' Setting Values
open_price = Range("C2").Value
summary_row = 2
Total_stock_volume = 0

' Setting Titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
' counting rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop and identify first row
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    
' Values for ticker, yearly change, and percent change
        ticker = ws.Cells(i, 1).Value
        closing_price = ws.Cells(i, 6).Value
        yearly_change = closign_price - open_price
        
        If open_price = 0 Then
     
        percent_change = 0
                
        Else
        
        percentage_change = yearly_change / open_price
        
        End If
        
' Sum of total stock volume
        Total_stock_volume = Total_stock_volume + ws.Cells(i, 7)
    
' Defining rows for ticker, yearly change

        ws.Cells(summary_row, 9).Value = ticker
        ws.Cells(summary_row, 10).Value = yearly_change
    
' Conditioning colors
    
        If ws.Cells(summary_row, 10).Value > 0 Then
            ws.Cells(summary_row, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(summary_row, 10).Value < 0 Then
            ws.Cells(summary_row, 10).Interior.ColorIndex = 3
        
        End If
    
' Defining percentage change and total stock volume
        ws.Cells(summary_row, 11).Value = percent_change
        ws.Cells(summary_row, 12).Value = Total_stock_volume
        summary_row = summary_row + 1
        open_price = ws.Cells(i + 1, 3).Value
        Total_stock_volume = 0
    
    Else
    
        Total_stock_volume = Total_stock_volume + ws.Cells(i, 7)
        
    End If
    
    Next i

' Reseting Total stock volume

        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    Next ws
    
    
End Sub
