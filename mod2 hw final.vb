
'Code was completed with assistance from chat gpt and bcs learning assistant.
Sub CalculateYearlyChange()
    Dim i As Long
    Dim change As Double
    Dim start As Long
    Dim percentchange As Double
    start = 2
    j = 0
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            change = Cells(i, 6) - Cells(start, 3)
            percentchange = change / Cells(start, 3)
            
            Range("M" & 2 + j).Value = change
            Range("N" & 2 + j).Value = percentchange
            
            start = i + 1
            j = j + 1
        End If
    Next i
End Sub




'Code taken from class material and learning assistant on bcs

Sub SubTotal()
    Dim i As Long
    Dim RowCount As Long
    Dim sum As Double
    Dim sumrow As Long
    sumrow = 2
    sum = 0
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        sum = sum + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Cells(sumrow, 12).Value = Cells(i, 1).Value
            Cells(sumrow, 15).Value = sum
            sumrow = sumrow + 1
            sum = 0
        End If
    Next i
    Cells(1, 12).Value = "ticker"
    Cells(1, 13).Value = "yearly change"
    Cells(1, 14).Value = "percent change"
    Cells(1, 15).Value = "total volume"
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"
    

    Cells(2, 18) = "%" & WorksheetFunction.Max(Range("N2:N3001" & RowCount)) * 100
    Cells(3, 18) = "%" & WorksheetFunction.Min(Range("N2:N3001" & RowCount)) * 100
    Cells(4, 18) = WorksheetFunction.Max(Range("O2:O3001" & RowCount))
    Cells(sumrow, 15).Value = WorksheetFunction.sum(Range("F:F"))
    
End Sub

Sub AddConditionalFormatting()
    Dim rng As Range
    Dim formatRange As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    
  
    Set rng = ActiveSheet.Range("M:M")
    
    
    
    Set formatRange = rng
    
    
    Set condition1 = formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
    With condition1
        .Interior.ColorIndex = 4
    End With
    
    
    Set condition2 = formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    With condition2
        .Interior.ColorIndex = 3
    End With
    'Code obtained with help from chat gpt
    
End Sub