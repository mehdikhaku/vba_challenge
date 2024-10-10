Option Explicit

Sub TickerAnalysis()

' Set initial variables
    Dim ws As Worksheet
    Dim Ticker_Symbol As String
    Dim lastRow As Long
    Dim i As Long
    Dim Summary_Table_Row As Integer
    Dim Starting_Price As Double
    Dim Ending_Price As Double
    Dim Ticker_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Volume As Double
    Dim find_value As Long
        
' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets

' Set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
' Keep track of the location for each ticker symbol in the summary table
    Summary_Table_Row = 2
  
' Find the last row with data in column A (tickers)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
' Get the starting price of the first ticker symbol
    Starting_Price = ws.Cells(2, 3).Value
    Total_Volume = 0 ' Initialize total volume to zero
    
' Loop through all ticker symbols
    For i = 2 To lastRow
  
' Add the volume from the current row to the total volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
  
' Check if we are still within the same ticker symbol, if we are not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
' Set the ticker symbol
    Ticker_Symbol = ws.Cells(i, 1).Value
    
' Get the ending price for the current ticker symbol from column F
    Ending_Price = ws.Cells(i, 6).Value
 
' Find the first non-zero starting value
    If Starting_Price = 0 Then
        Dim start As Long
        start = i
        For find_value = start To lastRow
    If ws.Cells(find_value, 3).Value <> 0 Then
        Starting_Price = ws.Cells(find_value, 3).Value
    Exit For
    End If
    Next find_value
End If
        
' Calculate the change in price
    Ticker_Change = Ending_Price - Starting_Price
    
' Calculate the percentage change
    If Starting_Price <> 0 Then
      Percentage_Change = (Ticker_Change / Starting_Price) * 100
    Else
' Avoid division by zero
      Percentage_Change = 0
    End If
            
' Print the ticker symbol in coumn I
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
    
' Print the change in price in column J
    ws.Range("J" & Summary_Table_Row).Value = Ticker_Change
           
' Print the percentage change in column K
    ws.Range("K" & Summary_Table_Row).Value = Percentage_Change / 100
                       
' Print the total volume in column L
    ws.Range("L" & Summary_Table_Row).Value = Total_Volume
    
' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
 
' Update the starting price for the next ticker symbol
    Starting_Price = ws.Cells(i + 1, 3).Value
    
' Reset the total volume for the next ticker symbol
    Total_Volume = 0
     
    End If

  Next i
        
' Define Variables
    Dim rng As Range
    Dim cell As Range

' Set the range of cells to be checked and formatted
    Set rng = ws.Range("J2:J" & Summary_Table_Row - 1)

' Clear any existing conditional formatting
    rng.FormatConditions.Delete

' Loop through each cell in the range
    For Each cell In rng
    If cell.Value > 0 Then

' Color the cell green for positive values
    cell.Interior.ColorIndex = 4 ' Green
    
    ElseIf cell.Value < 0 Then
' Color the cell red for negative values
    cell.Interior.ColorIndex = 3 ' Red
    
    Else
' No fill for zero values
    cell.Interior.ColorIndex = xlNone
    
    End If
    
Next cell

' Format column J to two decimal places
    ws.Range("J2:J" & Summary_Table_Row - 1).NumberFormat = "0.00"

' Format column K to display as percentage
    ws.Range("K2:K" & Summary_Table_Row - 1).NumberFormat = "0.00%"

' Format column L as a number with commas and no decimal points
    ws.Range("L2:L" & Summary_Table_Row - 1).NumberFormat = "#,##0"
       
' Define Variables
    Dim rowCount As Long
    Dim increase_number As Long, decrease_number As Long, volume_number As Long

' Take the max and min and place them in a separate part in the worksheet
    rowCount = Summary_Table_Row - 1
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCount))
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCount))
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

' Final ticker symbol for total, greatest % of increase and decrease, and average
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

    ws.Range("P2").Value = ws.Cells(increase_number + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(decrease_number + 1, 9).Value
    ws.Range("P4").Value = ws.Cells(volume_number + 1, 9).Value

' Format Q2 and Q3 as percentages
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

' Format Q4 as a number with commas and no decimal points
    ws.Range("Q4").NumberFormat = "#,##0"

' Bold the header row (row 1) for specific columns
    ws.Range("I1:Q1").Font.Bold = True
    ws.Range("O2:O4").Font.Bold = True
    
' Autofit specific columns to ensure data fits
    ws.Columns("I:I").AutoFit
    ws.Columns("J:J").AutoFit
    ws.Columns("K:K").AutoFit
    ws.Columns("L:L").AutoFit
    ws.Columns("O:O").AutoFit
    ws.Columns("P:P").AutoFit
    ws.Columns("Q:Q").AutoFit

' End of worksheet loop
    Next ws

End Sub

