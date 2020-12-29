Attribute VB_Name = "Module1"
Sub Stockmarket()
   
Dim ws As Worksheet

For Each ws In Worksheets
   
Dim tickername As String
Dim tickervol As Double
Dim ticker_summary As Long
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
        
ticker_summary = 2
tickervol = 0
open_price = ws.Cells(2, 3).Value
          
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
       
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For I = 2 To lastrow
           
            If ws.Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                      
              tickername = ws.Cells(I, 1).Value
              
              tickervol = tickervol + ws.Cells(I, 7).Value

              ws.Range("I" & ticker_summary).Value = tickername
              
              ws.Range("L" & ticker_summary).Value = tickervol

              close_price = ws.Cells(I, 6).Value
            
              yearly_change = (close_price - open_price)
                           
              ws.Range("J" & ticker_summary).Value = yearly_change
             
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
                End If

              ws.Range("K" & ticker_summary).Value = percent_change
              ws.Range("K" & ticker_summary).NumberFormat = "0.00%"
                 
              ticker_summary = ticker_summary + 1
              
              tickervol = 0

              open_price = ws.Cells(I + 1, 3)
            
            Else
              
              tickervol = tickervol + ws.Cells(I, 7).Value

            
            End If
        
        Next I

    
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
           
    For I = 2 To lastrow_summary_table
            If ws.Cells(I, 10).Value > 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(I, 10).Interior.ColorIndex = 3
            End If
            
    Next I

    Next ws
     
End Sub

