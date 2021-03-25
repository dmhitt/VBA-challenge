'Name: Dinnara Hitt
'VBA Homework

Sub stock_totals()

For Each ws In Worksheets

    'Dim worsheet_Name As String
    
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim sum_volume As LongLong
    Dim row As Integer
    Dim col As Integer
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim first_time As Boolean
    

    row = 2
       
    sum_volume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    'MsgBox ("lastrow = " + Str(lastrow))
    'MsgBox ("lastcolumn = " + Str(lastcolumn))
    
    
    
    col = lastcolumn + 3
    
    first_time = True
    
    ws.Cells(1, col).Value = "Ticker"
    ws.Cells(1, col + 1).Value = "Yearly Change"
    ws.Cells(1, col + 2).Value = "Pecent Change"
    ws.Cells(1, col + 3).Value = "Total Stock Volume"
    
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
            sum_volume = 0
            row = row + 1
            first_time = True
            

        End If

    Next i
    
Next ws


End Sub