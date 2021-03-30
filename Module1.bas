Attribute VB_Name = "Module1"
Sub Stocks():

    Dim currentWS As Worksheet
   
    
    For Each currentWS In Worksheets
    
        Dim ticker As String
       
        
        Dim total_ticker_vol As Double
        
        
        Dim open_price As Double
        Dim close_price As Double
        Dim delta_price As Double
        Dim delta_percent As Double
        
        Dim max_ticker As String
        Dim min_ticker As String
        Dim max_percent As Double
        Dim min_percent As Double
        Dim max_vol_name As String
        Dim max_vol As Double
        max_percent = 0
        min_percent = 1000000
        Dim lrow As Long
        
        Dim counter As Long
        counter = 2
        
        lrow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row
        
            currentWS.Range("I1").Value = "Ticker"
            currentWS.Range("J1").Value = "Yearly Change"
            currentWS.Range("K1").Value = "Percent Changed"
            currentWS.Range("L1").Value = "Total Stock Volume"
            currentWS.Range("O2").Value = "Greatest % Increase"
            currentWS.Range("O3").Value = "Greatest % Decrease"
            currentWS.Range("O4").Value = "Greatest Total Volume"
            currentWS.Range("P1").Value = "Ticker"
            currentWS.Range("Q1").Value = "Value"
      
        
        open_price = currentWS.Cells(2, 3).Value
        
        For i = 2 To lrow
        
            If currentWS.Cells(i, 1).Value <> currentWS.Cells(i + 1, 1).Value Then
                
                ticker = currentWS.Cells(i, 1).Value
                
                close_price = currentWS.Cells(i, 6).Value
                delta_price = close_price - open_price
                
                If open_price <> 0 Then
                
                    delta_percent = ((delta_price / open_price) * 100)
            
                End If
                
                total_ticker_vol = total_ticker_vol + currentWS.Cells(i, 7).Value
                
                currentWS.Range("I" & counter).Value = ticker
                currentWS.Range("J" & counter).Value = delta_price
                
                If delta_price > 0 Then
                    currentWS.Range("J" & counter).Interior.ColorIndex = 4
                ElseIf delta_price < 0 Then
                    currentWS.Range("J" & counter).Interior.ColorIndex = 3
                End If
                
                
                currentWS.Range("K" & counter).Value = (CStr(delta_percent) & "%")
                currentWS.Range("L" & counter).Value = total_ticker_vol
                
                
                If delta_percent > max_percent Then
                    max_percent = delta_percent
                    max_ticker = ticker
                ElseIf delta_percent < min_percent Then
                    min_percent = delta_percent
                    min_ticker = ticker
                
                End If
                
                If total_ticker_vol > max_vol Then
                    max_vol = total_ticker_vol
                    max_vol_name = ticker
                End If
                
                total_ticker_vol = 0
                delta_percent = 0
                counter = counter + 1
                delta_price = 0
                delta_percent = 0
                close_price = 0
                open_price = currentWS.Cells(i + 1, 3).Value
            Else
                total_ticker_vol = total_ticker_vol + currentWS.Cells(i, 7).Value
            End If
            
            
          
                
         Next i
         
        currentWS.Range("Q2").Value = (CStr(max_percent) & "%")
        currentWS.Range("Q3").Value = (CStr(min_percent) & "%")
        currentWS.Range("P2").Value = max_ticker
        currentWS.Range("P3").Value = min_ticker
        currentWS.Range("Q4").Value = max_vol
        currentWS.Range("P4").Value = max_vol_name
                
               
    Next currentWS
                    
                       
    
End Sub
