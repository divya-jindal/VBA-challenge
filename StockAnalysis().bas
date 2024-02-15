Attribute VB_Name = "Module1"
Sub StockAnalysis()
    'integer for amount of sheets - NumSheets
    'integer for number of row being output to
        
    Dim outputRowNum, inputRowNum As Integer
    Dim ws As Worksheet
    Dim ticker, tickerPart As String
    Dim openVal, closeVal, percentChange As Double
    Dim lastRowIndex, totalStockVolume As Double
    
    'for greatest percent increase, decrease, and best total stock volume
    Dim allstarTickers(0 To 2) As String ' Declare as array of strings
    Dim allstarValues(0 To 2) As Double
    
    For Each ws In ThisWorkbook.Sheets 'go through each sheet in the workbook
        outputRowNum = 2
        inputRowNum = 2
        ticker = ws.Cells(inputRowNum, 1) 'ticker = first ticker in worksheet
        lastRowIndex = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        openVal = ws.Cells(inputRowNum, 3)
        totalStockVolume = 0
        
        'to reset and for header printing / reference
        allstarTickers(0) = "Greatest Percent Increase"
        allstarTickers(1) = "Greatest Percent Decrease"
        allstarTickers(2) = "Greatest Stock Volume"
        allstarValues(0) = 0#
        allstarValues(1) = 0#
        allstarValues(2) = 0#
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        
        'row labels for 2nd tier analysis
        For i = 2 To 4
            ws.Cells(i, 15) = allstarTickers(i - 2)
        Next i
        
        For r = 2 To lastRowIndex 'for each row to number of rows
            tickerPart = ws.Cells(r, 1)
            If (ticker = tickerPart) Then
                totalStockVolume = ws.Cells(r, 7) + totalStockVolume
            End If
            
            If ticker <> tickerPart Or r = lastRowIndex Then 'else if tickerpart is different is different or r = lastRowIndex
                'get close value
                If (ticker <> tickerPart) Then
                    closeVal = ws.Cells(r - 1, 6)
                ElseIf (r = lastRowIndex) Then
                    closeVal = ws.Cells(r, 6)
                End If
                
                'print results
                ws.Cells(outputRowNum, 9).Value = ticker 'print ticker to outputRowNum
                                
                ws.Cells(outputRowNum, 10) = closeVal - openVal 'print yearly change
                If closeVal - openVal < 0 Then
                    ws.Cells(outputRowNum, 10).Interior.Color = RGB(255, 0, 0)
                ElseIf closeVal - openVal > 0 Then
                    ws.Cells(outputRowNum, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(outputRowNum, 10).Interior.Color = RGB(255, 255, 0)
                End If
                    
                percentChange = (closeVal / openVal) * 100 - 100
                percentChange = Format(percentChange, "0.00")
                ws.Cells(outputRowNum, 11) = Str(percentChange) + "%" 'print percent change from open to close to outputRowNum
                
                'check if percentChange is the greatest or not
                If percentChange > allstarValues(0) Then
                    allstarValues(0) = percentChange
                    allstarTickers(0) = ticker
                ElseIf percentChange < allstarValues(1) Then
                    allstarValues(1) = percentChange
                    allstarTickers(1) = ticker
                End If
                
                ws.Cells(outputRowNum, 12) = totalStockVolume 'print total stock volume to outputRowNum
                If totalStockVolume > allstarValues(2) Then 'check if stock volume is the best
                    allstarValues(2) = totalStockVolume
                    allstarTickers(2) = ticker
                End If
                
                outputRowNum = outputRowNum + 1 'move to next row for output
                ticker = tickerPart  'change ticker to new ticker
                openVal = ws.Cells(r, 3) 'set new openVal
                totalStockVolume = ws.Cells(r, 7) 'restart totalStockVolume
            End If
        Next r
        
        'print greatest increase, decrease, and stock volume
        For i = 0 To 2
            ws.Cells(2 + i, 16) = allstarTickers(i)
            ws.Cells(2 + i, 17) = allstarValues(i)
        Next i
        ws.Cells(3, 17) = Str(ws.Cells(3, 17)) + "%" 'add %symbol to the percentage
        
    Next ws
End Sub


