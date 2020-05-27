Attribute VB_Name = "Module1"
Sub VBA_Wall_Street_Easy():

' Declaring Variables
 
Dim total_stock_volume_stock_volume As Double
Dim ticker As String
Dim lastrow As Long
Dim j As Long


'Looping through Multiple Wroksheet
For Each ws In Worksheets

' Selecting the entire row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

' Creating Heading for 2 New Columns
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Total Stock Volume"

'Set Initial total_stock_volume to 0 and setting J to 2
 j = 2
 total_stock_volume = 0
 

'Creating for Loops to determing Ticker and Total Stock Volume
 For i = 2 To lastrow
     If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
         total_stock_volume = total_stock_volume + ws.Range("G" & i).Value

     Else
         ticker = ws.Range("A" & i).Value
         ws.Range("I" & j).Value = ticker
         ws.Range("J" & j).Value = total_stock_volume + Range("G" & i).Value
         'To add new Row and reset thetotal_stock_volume
         j = j + 1
         total_stock_volume = 0
         
     End If

 Next i
Next ws
End Sub

