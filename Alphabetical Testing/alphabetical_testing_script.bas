Attribute VB_Name = "Module1"
Sub alphabetical_testing():
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Changed"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Dim ticker_name As String

        Dim volume_total As Double
        volume_total = 0

        Dim stock_open As Double
        Dim stock_close As Double

        Dim yearly_diff As Double
        yearly_diff = 0
        Dim percent_change As Double
        percent_change = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Dim ticker_count As Long
        ticker_count = 2

        Dim greatest_increase As Double
        greatest_increase = 0
        Dim greatest_decrease As Double
        greatest_decrease = 0
        Dim greatest_volume As Double
        greatest_volume = 0
        Dim pc1 As String

            For I = 2 To last_row
                
                stock_open = ws.Cells(ticker_count, 3).Value
                If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                    ticker_name = ws.Cells(I, 1).Value
                    volume_total = volume_total + ws.Cells(I, 7).Value
                    stock_close = ws.Cells(I, 6).Value
                    yearly_diff = stock_close - stock_open
                        If stock_open = 0 Then
                            percent_change = 0
                        Else
                            percent_change = (yearly_diff) / stock_open
                            pc1 = FormatPercent(percent_change)
                        End If
                    ws.Range("I" & Summary_Table_Row).Value = ticker_name
                    ws.Range("J" & Summary_Table_Row).Value = yearly_diff
                    ws.Range("K" & Summary_Table_Row).Value = pc1
                    ws.Range("L" & Summary_Table_Row).Value = volume_total
                    Summary_Table_Row = Summary_Table_Row + 1
                    ticker_count = I + 1
                    volume_total = 0
                Else
                    volume_total = volume_total + ws.Cells(I, 7).Value
                    
                End If

            Next I
            last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For I = 2 To last_row

            If ws.Cells(I, 10) > 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(I, 10).Interior.ColorIndex = 3
            End If

        Next I

        Dim greatest_ticker_increase As Integer
        Dim greatest_ticker_decrease As Integer
        Dim greatest_ticker_volume As Integer
        
            greatest_increase = WorksheetFunction.Max(ws.Range("K:K"))
            ws.Cells(2, 15).Value = "Greatest % Increase"
            Set FoundCell = ws.Range("K:K").Find(What:=greatest_increase)
            greatest_ticker_increase = WorksheetFunction.Match(greatest_increase, ws.Range("K:K"), 0)
            ws.Cells(2, 16).Value = ws.Cells(greatest_ticker_increase, 9)
            pc1 = FormatPercent(greatest_increase)
            ws.Cells(2, 17).Value = pc1
      
            
            greatest_decrease = WorksheetFunction.Min(ws.Range("K:K"))
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            greatest_ticker_decrease = WorksheetFunction.Match(greatest_decrease, ws.Range("K:K"), 0)
            ws.Cells(3, 16).Value = ws.Cells(greatest_ticker_decrease, 9)
            pc1 = FormatPercent(greatest_decrease)
            ws.Cells(3, 17).Value = pc1
    
            greatest_volume = WorksheetFunction.Max(ws.Range("L:L"))
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            greatest_ticker_volume = WorksheetFunction.Match(greatest_volume, ws.Range("L:L"), 0)
            ws.Cells(4, 16).Value = ws.Cells(greatest_ticker_volume, 9)
            ws.Cells(4, 17).Value = greatest_volume
    Next ws
    
End Sub
