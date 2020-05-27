Attribute VB_Name = "Module2"
Sub VBA_Wall_Street_Moderate()

' Declaring Variables

Dim ticker As String
Dim total_volume As Double
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double
Dim yearly_change As Double
Dim row_number As Long
Dim lastrow As Long
Dim x As Long
Dim j As Long

'Looping through Multiple Wroksheet
For Each ws In Worksheets

' Selecting the entire row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
 
 ' Creating Heading
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"
 
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

         x = x + 1
         total_volume = 0
         row_number = i + 1
         
     End If
 Next i
 Next ws
End Sub
 




