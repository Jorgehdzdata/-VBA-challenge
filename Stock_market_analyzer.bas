Attribute VB_Name = "Module1"
Sub ticker()

Dim ticker As String
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume_ticker As String
Dim opening_p As Double
Dim closing As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_stock_volume As Double

' last row using the xlUp
Dim LastRow As Long
' loop over each worksheet in the workbook
For Each ws In Worksheets

' I am not sure if this is needed but according to some sources in OverStock
  ws.Activate
' MsgBox ActiveSheet.Name

' Find the last row of each worksheet
 LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

' This is so I dont have to do the same header for every column
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("M1").Value = "Open Price"
    ws.Range("N1").Value = "Closing Price"
    
    
' This is to reset the count for each ticker

    ticker = ""
    opening_p = 0
    yearly_change = 0
    percent_change = 0
    total_stock_volume = 0
    summary_table_row = 2
    
' Ticker loop
    For i = 2 To LastRow
' Current ticker
        ticker = Cells(i, 1).Value
        
' start of the year opening_p price for current ticker.
        If opening_p = 0 Then
            opening_p = Cells(i, 3).Value
            Range("M" & summary_table_row).Value = opening_p
        End If
 
' Sum total stock volume for current ticker.
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
' Check if the next row within the ticker column is different
        If Cells(i + 1, 1).Value <> ticker Then
' Increment the number of tickers when we get to a different ticker in the list.
           
            Range("I" & summary_table_row).Value = ticker
            
'Closing price for ticker
            closing = Cells(i - 1, 6)
            Range("N" & summary_table_row).Value = closing
            
' Get yearly change value
            yearly_change = closing - opening_p
            
' Add yearly change value to the appropriate cell in each worksheet.
            Range("J" & summary_table_row).Value = yearly_change
            
' If yearly change value is greater than 0, shade cell green.
            If yearly_change > 0 Then
              Range("J" & summary_table_row).Interior.ColorIndex = 4
' If yearly change value is less than 0, shade cell red.
            ElseIf yearly_change < 0 Then
              Range("J" & summary_table_row).Interior.ColorIndex = 3
' If yearly change value is 0, shade cell yellow.
            Else
               Range("J" & summary_table_row).Interior.ColorIndex = 6
            End If
            
            
' Calculate percent change for each ticker
            If opening_p = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_p)
            End If
            
            
' Format the percent_change value as a percent.
            Range("K" & summary_table_row).Value = Format(percent_change, "Percent")
            
' Set open price back to 0 when we get to a different ticker in the list.
            opening_p = 0
            closing = 0
            
' Add total stock volume value to the appropriate cell in each worksheet.
            Range("L" & summary_table_row).Value = total_stock_volume
            
' Set total stock volume back to 0 when we get to a different ticker in the list.
            total_stock_volume = 0
            summary_table_row = summary_table_row + 1
        End If
         
    Next i
    
'****************************BONUS ASSIGNMENT *******************************
' Headers so I do not have to add them manually to each sheet
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
    
    
' Get the last row using the same statement we used in section one
 LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
 

' Assigning values to bonus veriables
        greatest_percent_increase = Cells(2, 11).Value
        greatest_percent_increase_ticker = Cells(2, 9).Value
        greatest_percent_decrease = Cells(2, 11).Value
        greatest_percent_decrease_ticker = Cells(2, 9).Value
        greatest_stock_volume = Cells(2, 12).Value
        greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
' loop through new ticker list
    For b = 2 To LastRow
    
' Ticker with greatest percent increase.
        If Cells(b, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(b, 11).Value
            greatest_percent_increase_ticker = Cells(b, 9).Value
        End If
        
' Ticker with greatest percent decrease.
        If Cells(b, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(b, 11).Value
            greatest_percent_decrease_ticker = Cells(b, 9).Value
        End If
        
' Ticker with greatest stock volume.
        If Cells(b, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(b, 12).Value
            greatest_stock_volume_ticker = Cells(b, 9).Value
        End If
        
    Next b
    
' Add the values for the bonus section to each worksheet.
        Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
        Range("Q2").Value = Format(greatest_percent_increase, "Percent")
        Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
        Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
        Range("P4").Value = greatest_stock_volume_ticker
        Range("Q4").Value = greatest_stock_volume
    
Next ws

End Sub
