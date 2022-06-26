Attribute VB_Name = "Module1"
Sub tickerResult():


Dim openPrice, closePrice, yearChange, percentChange As Double
Dim maxpercent, minpercent, maxvolume As Double
Dim ws_count As Integer
 
For ws_count = 1 To Worksheets.Count
Worksheets(ws_count).Select
bonus_ticker = ""
 
 ' Similar to credit card exercise in lesson, I used For/If Loop to get total charge and ticker name.
  totalStock = 0
  
  'Also lastrow function is added to make sure Loop knows when to stop.
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  resultRow = 2
  
  'Title each result column before starting loop
  
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"
  
  
  ' Loop through rows in the column
  For i = 2 To lastrow
    
    ' Searches for when the value of the next cell is different than that of the current cell
    'if name of ticker is different on next row, we save name of ticker and print out total Stock.
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        openPrice = Cells(i, 3).Value
    End If
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        totalStock = totalStock + Range("G" & i).Value
        closePrice = Cells(i, 6).Value
        yearChange = closePrice - openPrice
        percentChange = yearChange / openPrice

        Range("I" & resultRow).Value = Range("A" & i).Value
        Range("J" & resultRow).Value = Format(yearChange, "0.00")
        Range("K" & resultRow).Value = Format(percentChange, "0.00%")
        Range("L" & resultRow).Value = totalStock
        
        If Range("J" & resultRow).Value >= 0 Then
            Cells(resultRow, 10).Interior.ColorIndex = 4
        Else
            Cells(resultRow, 10).Interior.ColorIndex = 3
        End If
         
        resultRow = resultRow + 1
        totalStock = 0
        openPrice = 0
        closePrice = 0
        
    Else
        totalStock = totalStock + Range("G" & i).Value

    End If
         
  Next i
  
'Autofit all result columns
Range("I1:O1").EntireColumn.AutoFit

'Bonus part
'Defining length of resultrow that was computed previously.
Bonusrow = Cells(Rows.Count, 9).End(xlUp).Row

'Compute maxpercent, minpercent, and maxvolume from Column I to Column J
maxpercent = WorksheetFunction.Max(Range("K2:K" & Bonusrow))
Range("Q2").Value = Format(maxpercent, "0.00%")

minpercent = WorksheetFunction.Min(Range("K2:K" & Bonusrow))
Range("Q3").Value = Format(minpercent, "0.00%")

maxvolume = WorksheetFunction.Max(Range("L2:K" & Bonusrow))
Range("Q4").Value = maxvolume

'Find Ticker name that correspond to maxpercent, minpercent, and maxvolume
'This For loop will go through Column K and Column L and find row number that corresponds to that value.
'After finding corresponding row number, Ticker name will be transfer to Column I.

For i = 2 To Bonusrow
    If Range("K" & i).Value = maxpercent Then
    Range("P2").Value = Range("I" & i).Value
    ElseIf Range("K" & i).Value = minpercent Then
    Range("P3").Value = Range("I" & i).Value
    ElseIf Range("L" & i).Value = maxvolume Then
    Range("P4").Value = Range("I" & i).Value
       
    End If
    Next i

Next ws_count

End Sub
