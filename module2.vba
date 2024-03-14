Sub Ticker()
    
     Dim Ticker As String ' Declase data as String
     Dim j As Integer ' Declare variable j
     Dim ws As Worksheet
     Set ws = ThisWorkbook.Worksheets("A") ' Change "Sheet1" to your sheet's name

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
     
     j = 2 ' Initialize j
     
     For i = 2 To lastRow 'Set range
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
         Ticker = Cells(i, 1).Value 'Read value
         Cells(j, 9).Value = Ticker 'Add value
         j = j + 1 ' Increment j for each iteration
    End If
    
     Next i
    
     
End Sub

Sub yearlychange()

    Dim opening As Double
    Dim closing As Double
    Dim j As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("A")
    Dim Ticker As String

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    j = 2

    For i = 2 To lastRow
        
       
        If Cells(i, 1) <> Cells(i - 1, 1) Then
            ' New stock ticker, reset variables
            Ticker = Cells(i, 1).Value
            opening = Cells(i, 3).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Last day of the year for this stock
            closing = Cells(i, 6).Value
         
        Cells(j, 10) = closing - opening
        j = j + 1
        End If
        Next i
    Cells(1, 10).Value = "Yearly Change"

End Sub

Sub percentagechange()

Dim yearlychange As Double
Dim opening As Double
Dim j As Integer
 Dim ws As Worksheet
 Dim Ticker As String
     Set ws = ThisWorkbook.Worksheets("A")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

j = 2

 For i = 2 To lastRow
 
    If Cells(i, 1) <> Cells(i - 1, 1) Then
            ' New stock ticker, reset variables
            Ticker = Cells(i, 1).Value
            opening = Cells(i, 3).Value
    End If
    yearlychange = Cells(i, 10).Value
    Cells(j, 11) = (yearlychange / opening) * 100 & "%"
    Cells(j, 11).NumberFormat = "0.00%"
    j = j + 1
    Next i
    Cells(1, 11).Value = "Percentage Change"
    
End Sub

Sub totalstockvolume()

    Dim totalvolume As Double
    Dim j As Long
    Dim ws As Worksheet
    Dim Ticker As String
        Set ws = ThisWorkbook.Worksheets("A")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    j = 2
   

     For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' If ticker changes, write totalvolume to the cell and reset totalvolume
            ws.Cells(j - 1, 12).Value = totalvolume
            j = j + 1
            totalvolume = ws.Cells(i, 7).Value
        Else
            ' If ticker remains the same, add volume to totalvolume
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    Cells(1, 12) = "Total Stock Volume"
    
End Sub

Sub Bonus()

Cells(2, 14).Value = "Great Percent Increase"
Cells(3, 14).Value = "Greatest Percent Decrease"
Cells(4, 14).Value = "Highest Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Amount"

    Dim totalvolume As Double
    Dim j As Long
    Dim ws As Worksheet
    Dim Ticker As String
        Set ws = ThisWorkbook.Worksheets("A")
    Dim increase As Double
    Dim decrease As Integer
    Dim highvolume As Long
    

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    j = 2
    
    increase = 0
   

     For i = 2 To lastRow
        If Cells(i, 11).Value > increase Then
        increase = Cells(i, 11).Value

     End If
     Next i
     Cells(2, 16) = increase
     Cells(2, 16).NumberFormat = "0.00%"

End Sub
