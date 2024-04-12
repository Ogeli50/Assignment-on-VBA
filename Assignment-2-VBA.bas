Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim summaryRow As Integer
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red
                End If
                
                summaryRow = summaryRow + 1
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub

Sub FindGreatestValues()
    ' Variables to track the greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim lastRow As Long
    Dim percentChange As Double

    ' Initialize variables
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through the data to find the greatest values
    For i = 2 To lastRow
        ' Calculate percentage change
        If Cells(i, 4).Value <> 0 Then
            percentChange = (Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value
        Else
            percentChange = 0
        End If

        ' Update greatest % increase and % decrease
        If percentChange > greatestIncrease Then
            greatestIncrease = percentChange
            greatestIncreaseTicker = Cells(i, 1).Value
        ElseIf percentChange < greatestDecrease Then
            greatestDecrease = percentChange
            greatestDecreaseTicker = Cells(i, 1).Value
        End If

        ' Update greatest total volume
        If Cells(i, 7).Value > greatestVolume Then
            greatestVolume = Cells(i, 7).Value
            greatestVolumeTicker = Cells(i, 1).Value
        End If
    Next i

    ' Print out the results
    MsgBox "Greatest % Increase: " & greatestIncreaseTicker & " (" & Format(greatestIncrease, "0.00%") & ")" & vbCrLf & _
           "Greatest % Decrease: " & greatestDecreaseTicker & " (" & Format(greatestDecrease, "0.00%") & ")" & vbCrLf & _
           "Greatest Total Volume: " & greatestVolumeTicker & " (" & greatestVolume & ")"
End Sub


