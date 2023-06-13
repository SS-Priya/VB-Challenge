VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alp_test()

'Looping thru worksheets

 For Each ws In Worksheets
 
    Dim lastrow, count, i As Long
     
    Dim year_change, open_v, close_v As Double
 
    Dim ts_volume, max_tsc As LongLong
    
    Dim max_tic, min_tic, max_tsc_tic As String
    
    Dim max_value, min_value As Double
 
 'Counting Row
 
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
      
    
 'Setting Row names

 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
   
 ' Setting intial values to variables
    
    max_value = 0
    min_value = 0
    max_tsc = 0
    count = 2
    open_v = ws.Cells(2, 3).Value


    
  'Loop to evaluate Ticker, Yearly Change, Percent Change and Total Stock Volume
  'greatest increase, greatest decrease and greatest total value
        
        For i = 2 To lastrow
        
        ' To set yearly change &  percent change
            
             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                close_v = ws.Cells(i, 6).Value
                
                year_change = close_v - open_v
                
                
                ws.Cells(count, 10).Value = year_change
                
    'Setting differnt cell color depending on positive and negative values
        
                        
               If ws.Cells(count, 10).Value < 0 Then

                    ws.Cells(count, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(count, 10).Interior.ColorIndex = 4

                End If
         
    'Finding Percent Change, Greatest % increase and Greatest % decrease
    
                ws.Cells(count, 11).Value = year_change / open_v
                
                If ws.Cells(count, 11).Value >= max_value Then
                    max_value = ws.Cells(count, 11).Value
                    max_tic = ws.Cells(count, 9).Value
                End If
                
                If ws.Cells(count, 11).Value <= min_value Then
                    min_value = ws.Cells(count, 11).Value
                    min_tic = ws.Cells(count, 9).Value
                End If
                
                
                ws.Cells(count, 11).NumberFormat = "0.00%"
         
                open_v = ws.Cells(i + 1, 3).Value
               
                ts_volume = 0
         
                count = count + 1
             
            End If
            
    'To set  ticker , total stock volume
    
             If i < lastrow Then
                
                ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
                ts_volume = ts_volume + ws.Cells(i, 7)
                ws.Cells(count, 12).Value = ts_volume
            
             End If
             
     'Findng Greatest Total Value
            
            If ts_volume > max_tsc Then
              
                max_tsc_tic = Cells(count, 9).Value
                max_tsc = ts_volume
            
            End If

      Next i
      
        ws.Range("p2").Value = max_tic
        ws.Range("q2").Value = max_value
        ws.Range("q2").NumberFormat = "0.00%"

        ws.Range("p3").Value = min_tic
        ws.Range("q3").Value = min_value
        ws.Range("q3").NumberFormat = "0.00%"

        ws.Range("p4").Value = max_tsc_tic
        ws.Range("q4").Value = max_tsc
      
    Next ws
      
      
End Sub



