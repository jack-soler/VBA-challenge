Sub RunAll()

    Call Ticker
    Call yearlychange
    Call percentagechange
    Call totalstockvolume
    Call Bonus
    Call highvolume

End Sub

Sub Ticker()
    
     Dim Ticker As String
     Dim j As Integer
     Dim ws As Worksheet
     Dim lastRow As Long
     
     For Each ws In ThisWorkbook.Worksheets
     
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
        j = 2 '
     
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(j, 9).Value = Ticker
                j = j + 1 '
            End If
        Next i
        
        ws.Cells(1, 9) = "Ticker"
    
    Next ws
     
End Sub

Sub yearlychange()

    Dim opening As Double
    Dim closing As Double
    Dim j As Integer
    Dim ws As Worksheet
    Dim Ticker As String
    Dim lastRow As Long
    
    For Each ws In ThisWorkbook.Worksheets
     
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        j = 2

        For i = 2 To lastRow
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
                Ticker = ws.Cells(i, 1).Value
                opening = ws.Cells(i, 3).Value
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                closing = ws.Cells(i, 6).Value
                ws.Cells(j, 10) = closing - opening
                j = j + 1
            End If
        Next i
        
        ws.Cells(1, 10).Value = "Yearly Change"
    
        For i = 2 To lastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
            End If
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
            End If
        Next i
    
    Next ws
            

End Sub

Sub percentagechange()

    Dim yearlychange As Double
    Dim opening As Double
    Dim j As Integer
    Dim ws As Worksheet
    Dim lastRow As Long

    For Each ws In ThisWorkbook.Worksheets
     
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        j = 2

        For i = 2 To lastRow
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
                opening = ws.Cells(i, 3).Value
            End If
            yearlychange = ws.Cells(i, 10).Value
            ws.Cells(j, 11) = (yearlychange / opening) * 100 & "%"
            ws.Cells(j, 11).NumberFormat = "0.00%"
            j = j + 1
        Next i
        ws.Cells(1, 11).Value = "Percentage Change"
    
    Next ws
    
End Sub

Sub totalstockvolume()

    Dim totalvolume As Double
    Dim j As Long
    Dim ws As Worksheet
    Dim lastRow As Long
      
    For Each ws In ThisWorkbook.Worksheets
     
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        j = 2
        totalvolume = 0

        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ws.Cells(j - 1, 12).Value = totalvolume
                j = j + 1
                totalvolume = ws.Cells(i, 7).Value
            Else
                totalvolume = totalvolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ws.Cells(1, 12) = "Total Stock Volume"

    Next ws
    
End Sub

Sub Bonus()

    Dim j As Long
    Dim ws As Worksheet
    Dim Tickerhigh As String
    Dim Tickerlow As String
    Dim increase As Double
    Dim decrease As Double
    Dim lastRow As Long

    For Each ws In ThisWorkbook.Worksheets
     
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        increase = 0
        decrease = 0

        For i = 2 To lastRow
            If ws.Cells(i, 11).Value > increase Then
                increase = ws.Cells(i, 11).Value
                Tickerhigh = ws.Cells(i, 9)
            End If
        Next i

        For i = 2 To lastRow
            If ws.Cells(i, 11).Value < decrease Then
                decrease = ws.Cells(i, 11).Value
                Tickerlow = ws.Cells(i, 9)
            End If
        Next i

        ws.Cells(2, 16) = increase
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(2, 15) = Tickerhigh
        ws.Cells(3, 16) = decrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15) = Tickerlow
        
        ws.Cells(2, 14).Value = "Greatest Percent Increase"
        ws.Cells(3, 14).Value = "Greatest Percent Decrease"
        ws.Cells(4, 14).Value = "Highest Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Amount"
    
    Next ws

End Sub

Sub highvolume()

    Dim highvolume As Double
    Dim j As Long
    Dim ws As Worksheet
    Dim Tickervolume As String
    Dim lastRow As Long

    For Each ws In ThisWorkbook.Worksheets
     
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        highvolume = 0

        For i = 2 To lastRow
            If ws.Cells(i, 12).Value > highvolume Then
                highvolume = ws.Cells(i, 12).Value
                Tickervolume = ws.Cells(i, 9)
            End If
        Next i
        
        ws.Cells(4, 16) = highvolume
        ws.Cells(4, 15) = Tickervolume

    Next ws
        
End Sub
