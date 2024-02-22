Sub Очистка_листа_объекты()
    a = MsgBox("Удалить все значения?", vbQuestion + vbYesNo + vbDefaultButton2)
    If a = vbNo Then GoTo exitSub

    Application.Calculation = xlCalculationManual
    Sheets("Объекты").Select
    lastrow = Cells(1, 1).CurrentRegion.Rows.Count
    lastcolumn = Cells(1, 1).CurrentRegion.Columns.Count
    For j = 5 To lastcolumn
        If InStr(1, Cells(1, j), "Итог") = 0 Then
                Range(Cells(3, j), Cells(lastrow, j)).ClearContents
        End If
    Next j

exitSub:
    Application.Calculation = xlCalculationAutomatic
End Sub