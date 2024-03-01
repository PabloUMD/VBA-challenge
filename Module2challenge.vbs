Sub Module2challenge():

 Dim ws As Worksheet
 Dim rng As Range
For Each ws In ThisWorkbook.Worksheets
ws.Columns("I:Z").Delete


' Set an initial variables
  Dim Ticker As String
  Dim Firstopen As Double: Firstopen = 0
  Dim Lastclose As Double: Lastclose = 0
  Dim maxVal As Double, minVal As Double, totalVal As Double
  Dim maxTicker As String, minTicker As String, totalticker As String
  
  'Print  Total Headers
   ws.Cells(1, "P").Value = "Ticker"
   ws.Cells(1, "Q").Value = "Value"
   ws.Cells(2, "O").Value = "Greatest%increase"
   ws.Cells(3, "O").Value = "Greatest%Decrease"
   ws.Cells(4, "O").Value = "Greatest Total Volume"

  ' Define initial variables
  Dim Total_Stock As Double: Total_Stock = 0
  maxVal = 0
  minVal = 0
  totalVal = 0
  maxTicker = ""
  minTicker = ""
  totalticker = ""
 
  'Set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Cells(1, 12).Value = "Total Stock Volume"
  
  ' Keep track of the location for Ticker in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  Firstopen = Cells(2, 3).Value
  
  ' Loop through all ticker symbol
  For i = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      
      Ticker = Cells(i, 1).Value
      
      ' Add to the Total_Stock

     Total_Stock = Total_Stock + Cells(i, 7).Value
     
     'Substract Close Last value to Open first value
      ws.Range("J" & Summary_Table_Row).Value = Lastclose - Firstopen
      
      'Calculate the Percent change
      If Lastclose <> 0 And Firstopen <> 0 Then
      ws.Range("K" & Summary_Table_Row).Value = ((Lastclose - Firstopen) / Firstopen)
      End If

      ' Print the Ticker in the Summary Table
     ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Stock

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset
      Total_Stock = 0
      Firstopen = Cells(i + 1, 3).Value
      Lastclose = 0
    
   
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Total Stock
      Total_Stock = Total_Stock + Cells(i, 7).Value
      Lastclose = Cells(i + 1, 6).Value
      
    End If
   
    

  Next i
  
  'Calulate Max and Min values
    For i = 2 To ws.Cells(Rows.Count, "I").End(xlUp).Row
            
        If ws.Cells(i, "K").Value > maxVal Then
            maxVal = ws.Cells(i, "K").Value
            maxTicker = ws.Cells(i, "I").Value
        End If
            
        If ws.Cells(i, "K").Value < minVal Then
            minVal = ws.Cells(i, "K").Value
            minTicker = ws.Cells(i, "I").Value
        End If
        
        If ws.Cells(i, "L").Value > totalVal Then
            totalVal = ws.Cells(i, "L").Value
            totalticker = ws.Cells(i, "I").Value
        End If
      
    Next i
    
    'Print max and min values
    
    ws.Cells(2, "P").Value = maxTicker
    ws.Cells(2, "Q").Value = maxVal
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Cells(3, "P").Value = minTicker
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Cells(3, "Q").Value = minVal
    ws.Cells(4, "P").Value = totalticker
    ws.Cells(4, "Q").Value = totalVal
    
    'Condiitonal format with colors
    
    Set rng = ws.Range("J2:J" & ws.Cells(Rows.Count, "J").End(xlUp).Row)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = 255
        .Font.Color = RGB(255, 255, 255)
        .SetFirstPriority
    End With
    
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .SetFirstPriority
    End With

   
  Next ws
  

End Sub

