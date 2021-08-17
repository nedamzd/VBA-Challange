Sub alphabetical_testing()
 
 
 
 Dim Ticker As String
 Dim total_stock As Double
 Dim yearly_change As Double
 Dim temp_open As Double
 Dim temp_close As Double
 Dim temp_initial_open As Double
 Dim percentage As Double
 
 
 
 
 
  For Each ws In Worksheets
   Summary_Table_Row = 2
   total_stock = 0
   
   
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   'Create new column headers in each worksheet
   
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = " Yearly Change"
    ws.Cells(1, 12).Value = "Percentage Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
   
   temp_open = ws.Cells(2, 3).Value
   
   For i = 2 To LastRow
        
         
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         temp_close = ws.Cells(i, 6).Value
         yearly_change = temp_close - temp_open
         yearly_change_1 = Round(yearly_change, 3)
       
           
           If temp_open > 0 Then
            percentage = yearly_change / temp_open
           
           Else
            percentage = 0
           
           End If
           
        
         temp_open = ws.Cells(i + 1, 3).Value
         Ticker = ws.Cells(i, 1).Value
         total_stock = total_stock + ws.Cells(i, 7).Value
    
        
        'Print ticker in summary table
        ws.Range("j" & Summary_Table_Row).Value = Ticker
        ws.Range("M" & Summary_Table_Row).Value = total_stock
        ws.Range("k" & Summary_Table_Row).Value = yearly_change_1
        ws.Range("L" & Summary_Table_Row).Value = percentage
        ws.Range("L" & Summary_Table_Row).Style = "Percent"
        
        ws.Columns("A:L").AutoFit
        
        If ws.Range("k" & Summary_Table_Row).Value > 0 Then
         ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
         Else
         ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the total stock for the new ticker
        total_stock = 0
    
        Else
        
        total_stock = total_stock + ws.Cells(i, 7).Value
      
        
        End If
        
    Next i
    
  Next ws
  
End Sub
