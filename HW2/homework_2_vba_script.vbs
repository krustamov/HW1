Sub stock_market_analysis()
Dim total As Double, i As Long, delta As Single, j As Integer, begin As Long, row_no As Long
Dim percent_delta As Single, dates As Integer, daily_delta As Single, average_delta As Single
'declare ws as Worksheet
Dim ws As Worksheet

'results can be printed to desired worksheet
Set wsr = Worksheets("results")

'setting up the initial values
j = 0
total = 0
delta = 0
begin = 2
daily_delta = 0

'getting the number of the last row with data
row_no = Cells(Rows.Count, "A").End(xlUp).row

For i = 2 To row_no
'printing results when a new ticker is discovered in the column

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
'storing results in variables
   
    total = total + Cells(i, 7).Value
    delta = (Cells(i, 6) - Cells(begin, 3))
    percent_delta = Round((delta / Cells(begin, 3) * 100), 2)
    daily_delta = daily_delta + (Cells(i, 4) - Cells(i, 5))
    
'calculating average change
    dates = (i - begin) + 1
    average_delta = daily_delta / dates
    
'begining of the next stock ticker
    begin = i + 1
    
'printing results
    wsr.Range("A1").Value = "Ticker"
    wsr.Range("B1").Value = "Change"
    wsr.Range("C1").Value = "% Change"
    wsr.Range("D1").Value = "Avg Change"
    wsr.Range("E1").Value = "Total"
    wsr.Range("A" & 2 + j).Value = Cells(i, 1).Value
    wsr.Range("B" & 2 + j).Value = Round(delta, 2)
    wsr.Range("C" & 2 + j).Value = "%" & percent_delta
    wsr.Range("D" & 2 + j).Value = average_delta
    wsr.Range("E" & 2 + j).Value = total
    
'setting up conditional coloring
    Select Case delta
        Case Is > 0
            wsr.Range("B" & 2 + j).Interior.ColorIndex = 4
        Case Is < 0
            wsr.Range("B" & 2 + j).Interior.ColorIndex = 3
        Case Else
            wsr.Range("B" & 2 + j).Interior.ColorIndex = 0
    End Select
    
'reset variables when a new ticker begins

    total = 0
    delta = 0
    j = j + 1
    dates = 0
    daily_delta = 0
    
'adding up values for each ticker

Else
    total = total + Cells(i, 7).Value
    delta = delta + (Cells(i, 6).Value - Cells(i, 3))
    
    'delta between high and low prices
    daily_delta = daily_delta + (Cells(i, 4) - Cells(i, 5))
    
    
    End If
Next i
    
   
    wsr.Range("I2").Value = "Ticker"
    wsr.Range("G3").Value = "Greatest Volume"
    wsr.Range("H3").Value = Application.WorksheetFunction.Max(wsr.Columns("E"))
    
      
    wsr.Range("G5").Value = "Greatest % Increase"
    wsr.Range("H5").Value = "%" & Round(Application.WorksheetFunction.Max(wsr.Columns("C")), 2) * 100
        
    
    wsr.Range("G7").Value = "Greatest % Decrease"
    wsr.Range("H7").Value = "%" & Round(Application.WorksheetFunction.Min(wsr.Columns("C")), 2) * 100
    
    wsr.Range("G9").Value = "Greatest Daily Avg. Change"
    wsr.Range("H9").Value = Application.WorksheetFunction.Max(wsr.Columns("D"))
    
End Sub