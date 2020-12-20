Sub Stock_Market()

'Make sure you know the TYPE of variable you want to assign to inform the interface
Dim ws As Worksheet
Dim Ticker_Column As String
Dim Year_Change As Double
Dim Percent_Change As Double
Dim Total_Vol As LongLong
Dim start_value As Double
Dim end_value As Double
Dim Summary_Row_Values As Integer

Year_Change = 0
Percent_Change = 0
Total_Vol = 0
start_value = 0
end_value = 0

'Set at 2 b/c the second row is where we want the loop to start inputting new values

start_row = 2

Summary_Row_Values = 2

'Begin the loop through all the worksheet
For Each ws In Worksheets

'Setting up the Ticker, Yearly Change, Percent Change, and Total Stock Volume column
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'Making the end row get counted upon possible adjustment
Dim End_Row As Long
End_Row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To End_Row


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        start_value = ws.Cells(start_row, 3)
        end_value = ws.Cells(i, 6)
        Total_Vol = Total_Vol + Cells(i, 7).Value
        Ticker_Name = ws.Cells(i, 1).Value
        Year_Change = end_value - start_value
        Percent_Change = Round((Year_Change / start_value) * 100, 2)
        
        
        
      ' Print values in the Summary Table
      ws.Range("I" & Summary_Row_Values).Value = Ticker_Name

      ws.Range("J" & Summary_Row_Values).Value = Year_Change
      
      ws.Range("K" & Summary_Row_Values).Value = Percent_Change
       
      ws.Range("L" & Summary_Row_Values).Value = Total_Vol

      ' Add one to the summary table row to keep the rows counting down
      Summary_Row_Values = Summary_Row_Values + 1
      
      ' Reset Values
      Total_Vol = 0
      Year_Change = 0
      Percent_Change = 0
      start_value = 0
      end_value = 0

    ' If the cell immediately following a row is the same brand...
    Else 'Make sure you know the TYPE of variable you want to assign to inform the interface you want to

      ' Add to the Brand Total
      Total_Vol = Total_Vol + Cells(i, 7).Value
     
End If

Next i

Next

Dim Last_Row As Long
Last_Row = Cells(Rows.Count, 10).End(xlUp).Row



For j = 2 To Last_Row

Summary_Year_Change = Cells(j, 10).Value

    If Summary_Year_Change < 0 Then
    
        Cells(j, 10).Interior.ColorIndex = 3
        
    Else
    
        Cells(j, 10).Interior.ColorIndex = 4
    
End If

Next j

End Sub

