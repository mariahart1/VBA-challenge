Attribute VB_Name = "Module1"
Sub Stocks()


Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim total_stock_volume As Double
Dim percent_change As Double
Dim start_data As Integer

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    start_data = 2
    previous_i = 1
    total_stock_volume = 0

    end_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To end_row

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value

            previous_i = previous_i + 1

            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value

            For j = previous_i To i
                total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value

            Next j

            If year_open = 0 Then
                percent_change = year_close
            Else
                yearly_change = year_close - year_open
                percent_change = yearly_change / year_open
            End If

            ws.Cells(start_data, 9).Value = ticker
            ws.Cells(start_data, 10).Value = yearly_change
            ws.Cells(start_data, 11).Value = percent_change

            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = total_stock_volume

            start_data = start_data + 1
            total_stock_volume = 0
            yearly_change = 0
            percent_change = 0
            previous_i = i

        End If

    Next i

    end_row_k = ws.Cells(Rows.Count, "K").End(xlUp).Row

    Increase = 0
    Decrease = 0
    Greatest = 0

        For k = 3 To end_row_k

            last_k = k - 1
            current_k = ws.Cells(k, 11).Value
            prevous_k = ws.Cells(last_k, 11).Value
            volume = ws.Cells(k, 12).Value
            prevous_vol = ws.Cells(last_k, 12).Value

            If Increase > current_k And Increase > prevous_k Then
                Increase = Increase
            ElseIf current_k > Increase And current_k > prevous_k Then
                Increase = current_k
                increase_name = ws.Cells(k, 9).Value
            ElseIf prevous_k > Increase And prevous_k > current_k Then
                Increase = prevous_k
                increase_name = ws.Cells(last_k, 9).Value
            End If

            If Decrease < current_k And Decrease < prevous_k Then
                Decrease = Decrease
            ElseIf current_k < Increase And current_k < prevous_k Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value
            ElseIf prevous_k < Increase And prevous_k < current_k Then
                Decrease = prevous_k
                decrease_name = ws.Cells(last_k, 9).Value
            End If
            
            If Greatest > volume And Greatest > prevous_vol Then
                Greatest = Greatest
            ElseIf volume > Greatest And volume > prevous_vol Then
                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value
            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                Greatest = prevous_vol
                greatest_name = ws.Cells(last_k, 9).Value
            End If
        Next k

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"

    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest
 
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"

    ws.Range("A1:P1").EntireColumn.AutoFit
    
    end_row_j = ws.Cells(Rows.Count, "J").End(xlUp).Row

        For j = 2 To end_row_j
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j
    
Next ws

End Sub
