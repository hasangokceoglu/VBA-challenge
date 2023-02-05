Attribute VB_Name = "Module1"
Sub multi_years()

Dim ws As Worksheet

For Each ws In Worksheets

Dim ticker_name As String
Dim lastrow As Double
Dim lastoutputrow As Double

Dim Max_Percentage As Double
Dim min_Percentage As Double
Dim Max_Total As Double


ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastoutputrow = 2

Dim yearly_change As Double
Dim opening_price As Double
Dim percentage_change As Double
Dim volume_total As Double

opening_price = ws.Cells(2, 3).Value
volume_total = 0


    For i = 2 To lastrow
    

          

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
            'ticker name assigment
        
            ticker_name = ws.Cells(i, 1).Value
        
            ws.Cells(lastoutputrow, 9).Value = ticker_name
        
        
            'yearly_change assigment
        
            yearly_change = ws.Cells(i, 6).Value - opening_price
        
            ws.Cells(lastoutputrow, 10).Value = yearly_change
          
        
            'percentage change assigment
        
            percentage_change = yearly_change / opening_price
        
            ws.Cells(lastoutputrow, 11).Value = percentage_change
        
            
            'update opening price
        
            opening_price = ws.Cells(i + 1, 3).Value
        
               
            'volume total calculation
        
            volume_total = volume_total + ws.Cells(i, 7).Value
        
            ws.Cells(lastoutputrow, 12).Value = volume_total
        
            volume_total = ws.Cells(i + 1, 9).Value
        
            'increase last output row
        
            lastoutputrow = lastoutputrow + 1
        
            
                Else
        
                volume_total = volume_total + ws.Cells(i, 7).Value
                       
    
        End If

    
    Next i



' min-max values
ws.Range("p2:p3").NumberFormat = "0.00%"
Max_Percentage = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(lastoutputrow - 1, 11)))
min_Percentage = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(lastoutputrow - 1, 11)))
Max_Total = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(lastoutputrow - 1, 12)))
ws.Cells(2, 16).Value = Max_Percentage
ws.Cells(3, 16).Value = min_Percentage
ws.Cells(4, 16).Value = Max_Total

        
        'summary ticker info
        
        For j = 2 To lastoutputrow - 1
        
            If ws.Cells(j, 11).Value = Max_Percentage Then
            ws.Cells(2, 15) = ws.Cells(j, 9)
            
            End If
            
            If ws.Cells(j, 11).Value = min_Percentage Then
            
            ws.Cells(3, 15) = ws.Cells(j, 9)
            
            End If
            
            If ws.Cells(j, 12).Value = Max_Total Then
            
            ws.Cells(4, 15) = ws.Cells(j, 9)
            
            End If
            
            'conditional formatting
            
            If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
            Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
            
            End If
            
          
        Next j
        

ws.Columns("K").NumberFormat = "0.00%"

ws.Columns("A:P").AutoFit
                    
Next ws


End Sub


