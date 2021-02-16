Sub Stock_Analysis()

'Instructions to run through each Worksheet *Keep Position*
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

' Set a variable for Ticker
  Dim Ticker As String
  
' Set all other variables
  Dim Ticker_Open As Double
  Dim Ticker_Close As Double
  Dim Ticker_High As Double
  Dim Ticker_Low As Double
  Dim Yearly_Change As Double
  Dim Percentage_Change As Double
  Dim Summary_Table_Row As Integer
  Dim i As Long
  
  Dim Ticker_Vol As Double
  Ticker_Vol = 0
   
   
'Set rows and columns for better comprehension when navigating *OPTIONAL*
  Dim Column As Integer
  Column = 1
  Dim row As Double
  row = 2

'Put Headers for analysis table
    Cells(1, Column + 10).Value = "Ticker"
    Cells(1, Column + 11).Value = "Yearly Change"
    Cells(1, Column + 12).Value = "Percent Change"
    Cells(1, Column + 13).Value = "Total Stock Volume"
 
 
'Use to find Last Row
     Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).row
 
  
'First Ticker_Open to avoid (For i) pick up
     Ticker_Open = Cells(2, 3).Value
  
' Loop through all tickers
     For i = 2 To Last_Row
     
  
' Check if we are still within the same ticker, if it is not Then...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
' Find all the Ticker names & Print into Summary
      Ticker = ws.Cells(i, Column).Value
      Cells(row, Collumn + 11).Value = Ticker
           
' Add to the Ticker_Close Total
      Ticker_Close = Cells(i, Collumn + 5).Value
      
' Calculate Yearly_Change & Print Yearly Change into the Summary Table
     Yearly_Change = Ticker_Close - Ticker_Open
     Cells(row, Column + 11).Value = Yearly_Change
  
' Calculate Percentage_Change
     
     If (Ticker_Open = 0 And Ticker_Close = 0) Then
                    Percentage_Change = 0
                    
     ElseIf (Ticker_Open = 0 And Ticker_Close <> 0) Then
                    Percentage_Change = 1
                    
     'Calculate the Percentage_Change and print in summarry
     Else
     Percentage_Change = Yearly_Change / Ticker_Open
                    Cells(row, Column + 12).Value = Percentage_Change
     
'Change the format of Percentage_Change to %
     Cells(row, Column + 12).NumberFormat = "0.00%"
    
End If
    
' Calculate the Ticker TotalVolumes & add Ticker Volume to the Summary Table

    Ticker_Vol = Ticker_Vol + Cells(i, Column + 6).Value
    Cells(row, Column + 13).Value = Ticker_Vol
                
' Add one to the summary table row
    row = row + 1

' Open Price for the rest of the table * move to here as advised by Kengo*
    Ticker_Open = Cells(i + 1, Column + 2)

' Reset the Volume Total
    Ticker_Vol = 0
        
'If cells have the same ticker
    Else
    Ticker_Vol = Ticker_Vol + Cells(i, Column + 6).Value
           
End If
     
'Finish the Loop
Next i

'*******************************************
' Conditional Formating loops
'Find the last row of Yearly_Change otherwise formatting will color the entire row
    Yearly_Change_Lastrow = ws.Cells(row, Column + 11).End(xlUp).row

For j = 2 To Yearly_Change_Lastrow

'Set Formating
    If (Cells(j, Column + 11).Value > 0 Or Cells(j, Column + 11).Value = "0") Then
    'Set to GREEN
    Cells(j, Column + 11).Interior.Color = vbGreen
    
    ElseIf Cells(j, Column + 11).Value < 0 Then
    
    'Set to RED
    Cells(j, Column + 11).Interior.Color = vbRed



End If

'Finish Loop
Next j

'*******************************************

' Set new headers
        Cells(2, Column + 16).Value = "Greatest % Increase"
        Cells(3, Column + 16).Value = "Greatest % Decrease"
        Cells(4, Column + 16).Value = "Greatest Total Volume"
        Cells(1, Column + 17).Value = "Ticker"
        Cells(1, Column + 18).Value = "Value"

' Loop through new rows to find the values and its associate ticker and print to new table
        For k = 2 To Yearly_Change_Lastrow
        
        'For Max
        If Cells(k, Column + 12).Value = WorksheetFunction.Max(Range("M2:M" & Yearly_Change_Lastrow)) Then
                'Ticker
                Cells(2, Column + 17).Value = Cells(k, Column + 10).Value
                'Value
                Cells(2, Column + 18).Value = Cells(k, Column + 12).Value
                Cells(2, Column + 18).NumberFormat = "0.00%"
        
        'For Min
        ElseIf Cells(k, Column + 12).Value = WorksheetFunction.Min(Range("M2:M" & Yearly_Change_Lastrow)) Then
                'Ticker
                Cells(3, Column + 17).Value = Cells(k, Column + 10).Value
                'Value
                Cells(3, Column + 18).Value = Cells(k, Column + 12).Value
                Cells(3, Column + 18).NumberFormat = "0.00%"
        
        'For Total_Vol
        ElseIf Cells(k, Column + 13).Value = WorksheetFunction.Max(Range("N2:N" & Yearly_Change_Lastrow)) Then
                'Ticker
                Cells(4, Column + 17).Value = Cells(k, Column + 10).Value
                'Value
                Cells(4, Column + 18).Value = Cells(k, Column + 13).Value
End If
        
Next k


'Finish the worksheet
Next ws
     

End Sub

