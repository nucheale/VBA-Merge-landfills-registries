Sub Изменить_даты_графиков()

    With Application
        .Calculation = xlCalculationManual
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With

    Set macroWb = ThisWorkbook
    macroWb.Worksheets("Графики").Activate

    calcMonth = InputBox("Введите номер месяца от 1 до 12, для которого нужно построить графики", "Изменение графиков", Month(Date))
    If Not IsNumeric(calcMonth) Then GoTo errorExit
    If CInt(calcMonth) < 1 Or CInt(calcMonth) > 12 Then GoTo errorExit
    calcMonth = CInt(calcMonth)
    firstDate = DateSerial(Year(Date), calcMonth, 1)
    If Not calcMonth = Month(Date) Then lastDate = DateSerial(Year(Date), calcMonth + 1, 1) - 1 Else lastDate = Date - 1
    
    With macroWb.Worksheets("Графики")
        chartTitlesTwoDim = Sheets("Справочник").ListObjects("LandfillsList").ListColumns("Для графиков").DataBodyRange.Value
        chartTitles = twoDimArrayToOneDim(chartTitlesTwoDim)
        For e = LBound(chartTitles) To UBound(chartTitles)
            For n = .ListObjects(chartTitles(e)).ListRows.Count To 2 Step -1 'удаляем все строки кроме первой
                .ListObjects(chartTitles(e)).ListRows(n).Delete
            Next n
            ' .ListObjects(chartTitles(e)).ListRows(1) = firstDate
            .ListObjects(chartTitles(e)).ListColumns("Дата").DataBodyRange.Cells(1).Value = firstDate
            For n = 1 To lastDate - firstDate + 1
                If Not n = 1 Then .ListObjects(chartTitles(e)).ListRows.Add
                .ListObjects(chartTitles(e)).ListColumns("Дата").DataBodyRange.Cells(n).Value = firstDate + n - 1
                ' .ListObjects(chartTitles(e)).ListRows(n) = firstDate + n - 1
            Next n
        Next e

        landfillNames = macroWb.Worksheets("Справочник").ListObjects("LandfillsList").ListColumns("Полигоны").DataBodyRange.Value
        lastRowCharts = .Cells.SpecialCells(xlLastCell).Row
        lastColumnCharts = .Cells.SpecialCells(xlLastCell).Column
        lastRowChart = .ListObjects(chartTitles(1)).ListRows.Count
        oneDayWidth = 50
        minChartWidth = 470
        For i = 1 To lastRowCharts 'перемещение графиков чтобы все красиво
            For e = LBound(landfillNames, 1) To UBound(landfillNames, 1)
                If .Cells(i, 1) = landfillNames(e, 1) Then
                ' If InStr(.Cells(i, 1), landfillNames(e, 1)) Then
                    .Cells(i, 2) = "Статистика ввоза на " & landfillNames(e, 1)
                    For Each chrt In .ChartObjects
                        If InStr(chrt.Chart.ChartTitle.Text, landfillNames(e, 1)) Then
                            chrt.Top = .Cells(i + 1, lastColumnCharts + 2).Top
                            chrt.Left = .Cells(i + 1, lastColumnCharts + 2).Left
                            chrt.Chart.ChartTitle.Text = "Статистика ввоза на " & landfillNames(e, 1)
                            chrt.Height = 510
                            chrt.Width = WorksheetFunction.Max(minChartWidth, lastRowChart * oneDayWidth)
                        End If
                    Next chrt
                End If
            Next e
        Next i
        
        
        '----------------удаление нулевых объектов из легенды-----------------------

        For Each chrt In .ChartObjects 'добавляем легенду заново
            chrt.Chart.SetElement (msoElementLegendBottom)
            chrt.Chart.Legend.Delete
            chrt.Chart.SetElement (msoElementLegendBottom)
        Next chrt

        For Each tbl In .ListObjects 'заголовки всех таблиц
            tblCounter = 1
            Dim headers As Variant
            ReDim headers(1 To tbl.ListColumns.Count, 1 To UBound(chartTitles))
            For i = LBound(headers, 1) To UBound(headers, 1)
                headers(i, tblCounter) = tbl.ListColumns(i).Name
            Next i
            tblCounter = tblCounter + 1
        Next tbl
        
        For Each chrt In .ChartObjects
            chrt.Chart.Legend.LegendEntries(UBound(headers, 1) - 1).Delete 'удаление Итого из легенды
        Next chrt
        
        For i = LBound(chartTitles) To UBound(chartTitles)
            For n = UBound(headers) - 1 To LBound(headers) + 1 Step -1
                tempSum = 0
                Set sumColumn = .ListObjects(i).ListColumns(n)
                For Each cell In sumColumn.DataBodyRange
                    If Not IsError(cell) Then tempSum = tempSum + CDbl(cell)
                Next cell
                If tempSum = 0 Then
                    For Each chrt In .ChartObjects
                        If InStr(chrt.Chart.ChartTitle.Text, landfillNames(i, 1)) Then 'находим нужный график, т.к. у графиков и таблиц не совпадают индексы
                            chrt.Chart.Legend.LegendEntries(n - 1).Delete
                        End If
                    Next chrt
                End If
            Next n
        Next i
    End With

errorExit:
    With Application
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

