Attribute VB_Name = "Module1"
Sub ticker()

For Each ws In Worksheets

Dim ticker_table_row As Integer
ticker_table_row = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly_Change"
ws.Range("K1").Value = "Percent_Change"
ws.Range("L1").Value = "Total_Volume"

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

Dim stock_name As String

Dim yearly_change As Double
yearly_change = 0

Dim total_vol As Double
total_vol = 0

Dim closing As Double
Dim opening As Double
Dim percent_change As Double

opening = ws.Cells(2, 3).Value

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    For I = 2 To lastrow
    
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                
            total_vol = total_vol + ws.Cells(I, 7).Value
                
            stock_name = ws.Cells(I, 1).Value
                        
            closing = ws.Cells(I, 6).Value
                
            yearly_change = closing - opening
                        
            percent_change = (yearly_change / opening)
            
            ws.Range("I" & ticker_table_row).Value = stock_name
                
            ws.Range("J" & ticker_table_row).Value = yearly_change
                
                If ws.Range("J" & ticker_table_row).Value <= 0 Then
                    ws.Range("J" & ticker_table_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & ticker_table_row).Interior.ColorIndex = 4
                End If
                
            ws.Range("K" & ticker_table_row).Value = percent_change
                
            ws.Range("K" & ticker_table_row).NumberFormat = "0.00%"
                
            ws.Range("L" & ticker_table_row).Value = total_vol
                
            ticker_table_row = ticker_table_row + 1
                    
            opening = ws.Cells(I + 1, 3).Value
            
            total_vol = 0
                
        Else
        
            total_vol = total_vol + (ws.Cells(I, 7).Value)
                
        End If
                
    Next I

' bonus ticker

Dim ticker As String
Dim maxpercent As Double
Dim minpercent As Double
Dim maxvolume As Double

maxpercent = WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("P2").Value = maxpercent

minpercent = WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("P3").Value = minpercent

maxvolume = WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("P4").Value = maxvolume

For j = 2 To lastrow

    If ws.Cells(j, 11).Value = maxpercent Then
        ticker = ws.Cells(j, 9).Value
        ws.Range("O2").Value = ticker
    End If
    If ws.Cells(j, 11).Value = minpercent Then
        ticker = ws.Cells(j, 9).Value
        ws.Range("O3").Value = ticker
    End If
    If ws.Cells(j, 12).Value = maxvolume Then
        ticker = ws.Cells(j, 9).Value
        ws.Range("O4").Value = ticker
    End If

Next j

Next ws

End Sub
