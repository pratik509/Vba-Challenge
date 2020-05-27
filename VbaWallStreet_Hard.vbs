Attribute VB_Name = "Module3"
Sub VBA_Wall_Street_Hard()

' Declaring Variables

Dim ticker As String
Dim total_volume As Double
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double
Dim yearly_change As Double
Dim row_number As Long
Dim lastrow As Long
Dim greatest_increase As Long
Dim ticker_greatest_increase As String
Dim greatest_decrease As Long
Dim ticker_greatest_decrease As String
Dim greatest_total_volume As Long
Dim ticker_greatest_total_volume As String
Dim lastrow_table As Long
Dim x As Long
Dim y As Long
Dim j As Long

'Looping through Multiple Wroksheet
For Each ws In Worksheets

' Selecting the entire row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
 
 ' Creating Heading
 ws.Range("I1").Value = "Ticker"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("L1").Value = "Total Stock Volume"
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"
 ws.Range("O2").Value = "Greatest % Increase"
 ws.Range("O3").Value = "Greatest % Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"
 
'Setting totals for the variables

total_volume = 0
row_number = 2
x = 2

'Creating for Loops to for each year calculations
For i = 2 To lastrow
     If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
         total_stock_volume = total_stock_volume + ws.Range("G" & i).Value

        Else
             ticker = ws.Range("A" & i).Value
            

'Year Change and Percent Change Calcutations
         close_price = ws.Range("F" & i)
         open_price = ws.Range("C" & row_number)
         yearly_change = close_price - open_price
         

 'Yearly Change and Percent Change into Display Cells
         ws.Range("K" & x).Value = percent_change
         ws.Range("L" & x).Value = total_volume + ws.Range("G" & i).Value
         ws.Range("J" & x).Value = yearly_change
         ws.Range("I" & x).Value = ticker
         ws.Range("K" & x).NumberFormat = "0.00%"

 
 'Precent Change Calulation
         If open_price = 0 Then
            percent_change = 0
         Else
            percent_change = yearly_change / open_price
         End If
 'Conditional Formating to display green as positive and red as negative
         If ws.Range("J" & x).Value > 0 Then
            ws.Range("J" & x).Interior.ColorIndex = 4
         Else
            ws.Range("J" & x).Interior.ColorIndex = 3
         End If

'Resets the total volume and open price and adds a new row into display cell for next ticker

         row_number = i + 1
         x = x + 1
         total_volume = 0
         
         
     End If
 Next i
 
'Find Greatest % Increase, Greatest % Decrease, Greatest Total Volume and Their Ticker
'Setting Initial Values and Cell refrences
 greatest_increase = ws.Range("K2" & 2).Value
 greatest_decrease = ws.Range("K2" & 2).Value
 greatest_total_volume = ws.Range("L2" & y).Value
 ticker_greatest_decrease = ws.Range("I2" & y).Value
 ticker_greatest_increase = ws.Range("I2" & y).Value
 ticker_greatest_total_volume = ws.Range("I2" & y).Value
 
 
 
 'Calculate Last Row Of Table Cells
 lastrow_table = ws.Cells(Rows.Count, "I").End(xlUp).row
 
 'Looping Through Each Row Of Table Cells to find Values
 For y = 2 To lastrow_table:
     If ws.Range("K" & y + 1).Value > greatest_increase Then
        greatest_increase = ws.Range("K" & y + 1).Value
        ticker_greatest_increase = ws.Range("I" & y + 1).Value
     ElseIf ws.Range("K" & y + 1).Value < greatest_decrease Then
        greatest_decrease = ws.Range("K" & y + 1).Value
        ticker_greatest_decrease = ws.Range("I" & y + 1).Value
     ElseIf ws.Range("L" & y + 1).Value > greatest_total_volume Then
        greatest_total_volume = ws.Range("L" & y + 1).Value
        ticker_greatest_total_volume = ws.Range("I" & y + 1).Value
     End If
 Next y

 'Displaying Greatest % Increase, Greatest % Decrease, Greatest Total Volume and Their Ticker into Display Cells
 ws.Range("P2").Value = ticker_greatest_increase
 ws.Range("P3").Value = ticker_greatest_decrease
 ws.Range("P4").Value = ticker_greatest_total_volume

 ws.Range("Q2").Value = greatest_increase
 ws.Range("Q3").Value = greatest_decrease
 ws.Range("Q4").Value = greatest_total_volume
 ws.Range("Q2:Q3").NumberFormat = "0.00%"
 
 Next ws



End Sub
 
