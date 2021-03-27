Sub stock_totals()

For Each ws In Worksheets

    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim sum_volume As LongLong
    Dim row As Integer
    Dim col As Integer
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim first_time As Boolean
    Dim lastrow As Long
    Dim great_increase As Double
    Dim great_decrease As Double
    Dim great_total_volume As LongLong
    Dim keep_ticker_inc As String
    Dim keep_ticker_dec As String
    Dim keep_ticker_vol As String
    
  
    row = 2
       
    sum_volume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    'lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    'MsgBox ("lastrow = " + Str(lastrow))
    'MsgBox ("lastcolumn = " + Str(lastcolumn))
    
    ws.Range("J1:R" & lastrow).Clear
        
    col = 10
    
    first_time = True
    
    ws.Cells(1, col).Value = "Ticker"
    ws.Cells(1, col + 1).Value = "Yearly Change"
    ws.Cells(1, col + 2).Value = "Pecent Change"
    ws.Cells(1, col + 3).Value = "Total Stock Volume"
    
    great_decrease = 0
    great_increase = 0
    great_total_volume = 0
    
    For i = 2 To lastrow
    
        If first_time = True Then
            ticker = ws.Cells(i, 1).Value
            open_price = ws.Cells(i, 3).Value
            first_time = False
            'MsgBox ("open_price = " + Str(open_price))
        End If
        
        sum_volume = sum_volume + ws.Cells(i, 7)
        
        If ticker <> ws.Cells(i + 1, 1).Value Or i = lastrow Then
            close_price = ws.Cells(i, 6)
            'MsgBox ("i = " + Str(i))
            'MsgBox ("close_price = " + Str(close_price))
            ws.Cells(row, col).Value = ticker
            yearly_change = close_price - open_price
            
            ws.Cells(row, col + 1).Value = yearly_change
            
            'Add Color to the cells
            If yearly_change < 0 Then
                ws.Cells(row, col + 1).Interior.ColorIndex = 3
                
            Else
                ws.Cells(row, col + 1).Interior.ColorIndex = 4
            End If
            
            
            If open_price <> 0 Then
                percent_change = Round((yearly_change / open_price) * 100, 2)
            Else
                percent_change = Round(yearly_change * 100, 2)
            End If
            
            ws.Cells(row, col + 2).Value = (Str(percent_change) + "%")
            ws.Cells(row, col + 3).Value = sum_volume
            
            If great_decrease = 0 Then
                keep_ticker_dec = ticker
                great_decrease = percent_change
            End If
            
            
            If great_decrease > percent_change Then
                keep_ticker_dec = ticker
                great_decrease = percent_change
            End If
            
             If great_increase = 0 Then
                keep_ticker_inc = ticker
                great_increase = percent_change
            End If
            
            
            If great_increase < percent_change Then
                keep_ticker_inc = ticker
                great_increase = percent_change
            End If
            
            if great_total_volume = 0 Then
                keep_ticker_vol = ticker
                great_total_volume = sum_volume
            End If

            If great_total_volume < sum_volume Then
                keep_ticker_vol = ticker
                great_total_volume = sum_volume
            End If

            sum_volume = 0
            row = row + 1
            first_time = True
            

        End If

    Next i


    ws.cells(2,col + 6).Value = "Greatest % Increase"
    ws.cells(3,col + 6).Value = "Greatest % Decrease"
    ws.cells(4,col + 6).Value = "Greatest Total Volume"
    ws.cells(1,col+7).Value = "TIcker"
    ws.cells(1,col+8).Value = "Value"
    ws.cells(2,col+7).Value = keep_ticker_inc
    ws.cells(2,col+8).Value = (str(great_increase)+"%")
    ws.cells(3,col+7).Value = keep_ticker_dec
    ws.cells(3,col+8).Value = (str(great_decrease)+"%")
    ws.cells(4,col+7).Value = keep_ticker_vol
    ws.cells(4,col+8).Value = great_total_volume


    
Next ws


End Sub


