'Attribute VB_Name = "Module2"
Sub Ticker():

Dim Ticker_Name As String

Dim Ticker_Total As Double
Ticker_Total = 0

Dim Sum_Table As Integer
Sum_Table_Row = 2
For n = 9 To 14

    If n = 9 Then
        Cells(1, n).Value = "Ticker"
    ElseIf n = 10 Then
        Cells(1, n).Value = "Yearly Change"
    ElseIf n = 11 Then
        Cells(1, n).Value = "Percent Change"
    ElseIf n = 12 Then
        Cells(1, n).Value = "Total Stock Volume"
    ElseIf n = 13 Then
        Cells(1, n).Value = "Year Open"
    Else
        Cells(1, n).Value = "Year Close"
    End If
Next n

    
NumRows = Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count
For i = 2 To NumRows
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_Name = Cells(i, 1).Value
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        Range("I" & Sum_Table_Row).Value = Ticker_Name
        Range("L" & Sum_Table_Row).Value = Ticker_Total
        Sum_Table_Row = Sum_Table_Row + 1
        Ticker_Total = 0
    Else
        Ticker_Total = Ticker_Total + Cells(i + 1, 7).Value
    End If
Next i

Dim Start_Year, End_Year, Perc_Year, End_Row, Start_Row As Integer

Start_Year = 0
End_Year = 0
Perc_Year = 0
End_Row = 2
Start_Row = 2

For j = 2 To NumRows
    If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
        Start_Year = Cells(j, 3).Value
        Cells(Start_Row, 13) = Start_Year
        Start_Row = Start_Row + 1
    Else
        Start_Year = 0
    End If
Next j

        
        
For k = 2 To NumRows
    If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
        End_Year = Cells(k, 6).Value
        Cells(End_Row, 14).Value = End_Year
        End_Row = End_Row + 1
    Else
        End_Year = 0
    End If
Next k
        
        
NumRows2 = Range("I:I").Cells.SpecialCells(xlCellTypeConstants).Count
For p = 2 To NumRows2
    Cells(p, 10).Value = (Cells(p, 13).Value - Cells(p, 14).Value)
        If Cells(p, 10).Value > 0 Then
            Cells(p, 10).Interior.ColorIndex = 4
        Else
            Cells(p, 10).Interior.ColorIndex = 3
        End If
        
Next p

For y = 2 To NumRows2
    Cells(y, 11).Value = ((Cells(y, 10).Value / Cells(y, 13).Value) * 100)
Next y


End Sub



