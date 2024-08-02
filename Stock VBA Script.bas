Attribute VB_Name = "Module1"
Sub Stocks()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, summaryRow As Long
    Dim ticker As String, openPrice As Double, closePrice As Double
    Dim totalVolume As Double, quarterlyChange As Double, percentChange As Double
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim greatestIncreaseTicker As String, greatestDecreaseTicker As String, greatestVolumeTicker As String
    
    ' Initialize greatest values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        totalVolume = 0
        
        ' Set up summary headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 16).Value = "Ticker"
        ws.Cells(3, 16).Value = "Greatest % Increase"
        ws.Cells(4, 16).Value = "Greatest % Decrease"
        ws.Cells(5, 16).Value = "Greatest Total Volume"
        
        ws.Cells(2, 17).Value = "Value"
        ws.Cells(3, 17).Value = "Value"
        ws.Cells(4, 17).Value = "Value"
        ws.Cells(5, 17).Value = "Value"
        
        ' Initialize starting open price
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through data
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
                ' Calculate values for current ticker
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                quarterlyChange = closePrice - openPrice
                
                ' Avoid division by zero
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice)
                Else
                    percentChange = 0
                End If
                
                ' Output results
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Conditional formatting
                If quarterlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Conditional formatting for percent change
                If percentChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0)
                ElseIf percentChange < 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Update greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                ' Reset for next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                If i < lastRow Then
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Output greatest values
        ws.Cells(2, 17).Value = greatestIncreaseTicker
        ws.Cells(2, 18).Value = greatestIncrease
        
        ws.Cells(3, 17).Value = greatestDecreaseTicker
        ws.Cells(3, 18).Value = greatestDecrease
        
        ws.Cells(4, 17).Value = greatestVolumeTicker
        ws.Cells(4, 18).Value = greatestVolume
        
        ' Format percentage and scientific notation
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("R").NumberFormat = "0.00%"
        ws.Cells(4, 18).NumberFormat = "0.00E+00"
        
        ' Autofit columns for better readability
        ws.Columns("I:Q").AutoFit
    Next ws
    
    MsgBox "Analysis complete!", vbInformation
End Sub

