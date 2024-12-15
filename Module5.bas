Attribute VB_Name = "Module5"
Sub AnalyzeStockData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    
    ' Loop through each worksheet (assuming each worksheet represents a stock)
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            closingPrice = ws.Cells(i, 6).Value
            totalVolume = ws.Cells(i, 7).Value
            
            ' Calculate quarterly change and percentage change
            quarterlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentageChange = (quarterlyChange / openingPrice) * 100
            Else
                percentageChange = 0
            End If
            
               ' Output the results
            ws.Cells(i, 11).Value = quarterlyChange ' Output quarterly change in column E
            ws.Cells(i, 12).Value = percentageChange ' Output percentage change in column F
            ws.Cells(i, 13).Value = totalVolume ' Output total volume in column G
        Next i
    Next ws
    
    
    ' Calculate percentage increase and decrease
        If openingPrice <> 0 Then
            percentIncrease = ((closingPrice - openingPrice) / openingPrice) * 100
            percentDecrease = ((openingPrice - closingPrice) / openingPrice) * 100
        Else
            percentIncrease = 0
            percentDecrease = 0
        End If

        ' Check for greatest percentage increase
        If percentIncrease > maxIncrease Then
            maxIncrease = percentIncrease
            stockIncrease = ws.Cells(i, 1).Value
        End If

        ' Check for greatest percentage decrease
        If percentDecrease > maxDecrease Then
            maxDecrease = percentDecrease
            stockDecrease = ws.Cells(i, 1).Value
        End If

        ' Check for greatest volume
        If volume > maxVolume Then
            maxVolume = volume
            stockVolume = ws.Cells(i, 1).Value
        End If
    Next i

    ' Output results
    MsgBox "Greatest % Increase: " & stockIncrease & " with " & maxIncrease & "%"
    MsgBox "Greatest % Decrease: " & stockDecrease & " with " & maxDecrease & "%"
    MsgBox "Greatest Total Volume: " & stockVolume & " with volume " & maxVolume
    
  

End Sub
