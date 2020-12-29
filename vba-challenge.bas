Attribute VB_Name = "Module1"
Sub Stockmarket()

            
        Dim tickername As String
        
        Dim tickervol As Double
        tickervol = 0
        
        Dim ticker_summary As Integer
        ticker_summary = 2
        
        Dim open_price As Double
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
       
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        For I = 2 To lastrow
           
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                      
              tickername = Cells(I, 1).Value
              
              tickervol = tickervol + Cells(I, 7).Value

              Range("I" & ticker_summary).Value = tickername
              
              Range("L" & ticker_summary).Value = tickervol

              close_price = Cells(I, 6).Value
            
              yearly_change = (close_price - open_price)
                           
              Range("J" & ticker_summary).Value = yearly_change
             
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
                End If

              Range("K" & ticker_summary).Value = percent_change
              Range("K" & ticker_summary).NumberFormat = "0.00%"
                 
              ticker_summary = ticker_summary + 1
              
              tickervol = 0

              open_price = Cells(I + 1, 3)
            
            Else
              
              tickervol = tickervol + Cells(I, 7).Value

            
            End If
        
        Next I

    
    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
       
    For I = 2 To lastrow_summary_table
            If Cells(I, 10).Value > 0 Then
                Cells(I, 10).Interior.ColorIndex = 4
        Else
                Cells(I, 10).Interior.ColorIndex = 3
        End If
            
    Next I

Next ws
    
End Sub
