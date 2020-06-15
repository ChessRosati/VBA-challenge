Attribute VB_Name = "Module1"
Sub wallStreet()

Dim Rows As Long
Dim totalVolume As Double
totalVolume = 0
Dim stockName As String
Dim summaryTableRow As Integer
Dim startingValue As Double
Dim endingValue As Double
Dim percentBasis As Double
Dim percent As String

startingValue = Cells(2, 3).Value
summaryTableRow = 2


With ActiveSheet
Rows = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

For i = 2 To Rows
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        stockName = Cells(i, 1).Value
        totalVolume = totalVolume + Cells(i, 7).Value
        endingValue = Cells(i, 6).Value
        Range("I" & summaryTableRow).Value = stockName
        Range("J" & summaryTableRow).Value = endingValue - startingValue
        If startingValue <> 0 Then
            percentBasis = Range("J" & summaryTableRow).Value / startingValue
            percent = FormatPercent(percentBasis)
            Range("K" & summaryTableRow).Value = percent
        Else
            Range("K" & summaryTableRow).Value = "n/a"
        End If
        Range("L" & summaryTableRow).Value = totalVolume
        
        summaryTableRow = summaryTableRow + 1
        totalVolume = 0
        startingValue = Cells(i + 1, 3).Value
    Else
        totalVolume = totalVolume + Cells(i, 7).Value
    End If
    Next i
End Sub
