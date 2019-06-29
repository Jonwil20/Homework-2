Attribute VB_Name = "Module1"
Sub Total_Stock_Volume()
Dim Ticker As String

Dim vol As Double
vol = 0

Dim summary_table_row As Integer
summary_table_row = 2
    
For i = 2 To 797711
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
    
    vol = vol + Cells(i, 7).Value
    
Range("I" & summary_table_row).Value = Ticker
Range("J" & summary_table_row).Value = vol

summary_table_row = summary_table_row + 1
vol = 0

Else
vol = vol + Cells(i, 7).Value
      
    End If

Next i
    
End Sub

