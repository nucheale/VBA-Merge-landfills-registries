Function getMaxTwoDArrayValue(arr) As Double
    maxValue = arr(LBound(arr), 1)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) > maxValue Then maxValue = arr(i, 1)
    Next i
    getMaxTwoDArrayValue = maxValue
End Function

Function getMinTwoDArrayValue(arr) As Double
    minValue = arr(LBound(arr), 1)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 1) < minValue Then minValue = arr(i, 1)
    Next i
    getMinTwoDArrayValue = minValue
End Function

Function twoDimArrayToOneDim(oldArr)
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For i = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(i) = oldArr(i, 1)
    Next i
    twoDimArrayToOneDim = newArr
End Function

Sub Загрузить_данные()

    Dim e, element, i, j, fileIndex As Long
    
    Set macroWb = ActiveWorkbook
    
    filesToOpen = Application.GetOpenFilename(FileFilter:="All files (*.*), *.*", MultiSelect:=True, Title:="Выберите файлы")
    If TypeName(filesToOpen) = "Boolean" Then Exit Sub
    
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False

    
    With macroWb.Worksheets("Справочник")
        landfillsCount = .ListObjects("LandfillsList").ListRows.Count
        Dim objects, landfillTitles, weight1tTitles, weight2Titles As Variant
        objects = .ListObjects("objects").DataBodyRange.Value 'названия объектов, два столбца
        landfillTitles = .ListObjects("titles").ListColumns("Полигон").DataBodyRange.Value 'Range("objects[Полигон]").Value 'названия ячейки где указывают полигон
        weight1tTitles = .ListObjects("titles").ListColumns("Вес на погрузке").DataBodyRange.Value 'названия ячейки где указывают вес на погрузке (на МСС/МПС)
        weight2Titles = .ListObjects("titles").ListColumns("Вес на полигоне").DataBodyRange.Value 'названия ячейки где указывают вес на разгрузке (вес полигона)
        
    '    Dim arrNames As Variant
    '   arrNames = Array(objects, landfillTitles, weight1tTitles, weight2Titles)
    '    For Each arr In arrNames
    '        For e = 1 To UBound(arr)
    '            If arr(e, 1) = "" Then arr(e, 1) = "NONE"
    '        Next e
    '    Next
        
        maxIndex = Application.Max(UBound(objects), UBound(landfillTitles), UBound(weight1tTitles), UBound(weight2Titles))
        On Error Resume Next
        For e = 1 To maxIndex
            If objects(e, 1) = "" Then objects(e, 1) = "NONE"
            If landfillTitles(e, 1) = "" Then landfillTitles(e, 1) = "NONE"
            If weight1tTitles(e, 1) = "" Then weight1tTitles(e, 1) = "NONE"
            If weight2Titles(e, 1) = "" Then weight2Titles(e, 1) = "NONE"
        Next e
        On Error GoTo 0
        'Debug.Print "элемент: ", landfillTitles(UBound(landfillTitles), 1)
        
    End With

    fileIndex = 1
    For Each file In filesToOpen
        Set objectWb = Application.Workbooks.Open(Filename:=filesToOpen(fileIndex))
    
            currentObject = Empty
            For e = 1 To UBound(objects) 'текущий объект по названию файла
                If InStr(LCase(objectWb.Name), LCase(objects(e, 2))) Then currentObject = objects(e, 1)
                If Not currentObject = Empty Then Exit For
            Next e
            ' Debug.Print currentObject
            If currentObject = Empty Then
                MsgBox "Название объекта не обнаружено в справочнике. Проверьте название файла: " & objectWb.Name
                GoTo errorExit
            End If
    
            Sort = False
            If InStr(LCase(objectWb.Name), "обработка") Then Sort = True 'определение МСС/МПС
            'Debug.Print Sort
        
    
        With objectWb.Worksheets("Вывоз")
            lastColumnObject = .Cells(1, 1).CurrentRegion.Columns.Count
    
            For j = 1 To lastColumnObject
                For e = LBound(landfillTitles) To UBound(landfillTitles)
                    If LCase(.Cells(1, j)) = LCase(landfillTitles(e, 1)) Then landfillTitleColumn = j
                Next e
                For e = LBound(weight1tTitles) To UBound(weight1tTitles)
                    If LCase(.Cells(1, j)) = LCase(weight1tTitles(e, 1)) Then weightObjectTitleColumn = j
                Next e
                For e = LBound(weight2Titles) To UBound(weight2Titles)
                    If LCase(.Cells(1, j)) = LCase(weight2Titles(e, 1)) Then weightLandfillTitleColumn = j
                Next e
            Next j
            
            If landfillTitleColumn = Empty Then
                MsgBox "В файле " & objectWb.Name & " обнаружен заголовок столбца Полигон, которого нет в справочнике"
                GoTo errorExit
            End If
            If weightObjectTitleColumn = Empty Then
                MsgBox "В файле " & objectWb.Name & " обнаружен заголовок столбца Вес объекта, которого нет в справочнике"
                GoTo errorExit
            End If
            If weightLandfillTitleColumn = Empty Then
                MsgBox "В файле " & objectWb.Name & " обнаружен заголовок столбца Вес объекта, которого нет в справочнике"
                GoTo errorExit
            End If

            lastRowObject = .Cells(Rows.Count, weightObjectTitleColumn).End(xlUp).Row

    
            Dim datesOfObject, landfillsOfObject, weights1Object, weights2Object As Variant 'даты объекта, полигоны объекта, веса объекта
            If lastRowObject <= 2 Then 'если только 1 рейс, то создается не массив, а просто переменная. нужно создавать массив, решение:
                datesOfObject = .Range(.Cells(2, 1), .Cells(2, 1)).Resize(1, 2).Value
                landfillsOfObject = .Range(.Cells(2, landfillTitleColumn), .Cells(2, landfillTitleColumn)).Resize(1, 2).Value
                weights1Object = .Range(.Cells(2, weightObjectTitleColumn), .Cells(2, weightObjectTitleColumn)).Resize(1, 2).Value
                weights2Object = .Range(.Cells(2, weightLandfillTitleColumn), .Cells(2, weightLandfillTitleColumn)).Resize(1, 2).Value
            Else
                datesOfObject = .Range(.Cells(2, 1), .Cells(lastRowObject, 1))
                landfillsOfObject = .Range(.Cells(2, landfillTitleColumn), .Cells(lastRowObject, landfillTitleColumn))
                weights1Object = .Range(.Cells(2, weightObjectTitleColumn), .Cells(lastRowObject, weightObjectTitleColumn))
                weights2Object = .Range(.Cells(2, weightLandfillTitleColumn), .Cells(lastRowObject, weightLandfillTitleColumn))
            End If

            'debug.print "datesOfObject: " & ubound(datesOfObject), "landfillsOfObject: " & ubound(landfillsOfObject), "weights1Object: " & ubound(weights1Object), "weights2Object: " & ubound(weights2Object),
            
            For e = LBound(datesOfObject, 1) To UBound(datesOfObject, 1) 'перевод дат в формат даты
                If IsDate(datesOfObject(e, 1)) Then datesOfObject(e, 1) = CDate(datesOfObject(e, 1))
            Next e
            
            minFileDate = CDate(getMinTwoDArrayValue(datesOfObject))
            maxFileDate = CDate(getMaxTwoDArrayValue(datesOfObject))
            If fileIndex = 1 Then lastDateTable = maxFileDate
            If maxFileDate > lastDateTable Then lastDateTable = maxFileDate 'максимальная дата, чтобы понять надо ли к графикам добавлять строку с новым днем или нет
            If lastDateTable > Date Then
                MsgBox "В файле " & objectWb.Name & " обнаружены данные за будущие даты (" & lastDateTable & ")"
                GoTo errorExit
            End If

            For e = LBound(weights1Object) To UBound(weights1Object) 'перевод кг в т
                If weights1Object(e, 1) < 0 Then
                    MsgBox "Обнаружен вес с отрицательным значением (" & weights1Object(e, 1) & "). Номер строки: " & e + 1
                    GoTo errorExit
                End If
                If weights1Object(e, 1) > 100 Then weights1Object(e, 1) = weights1Object(e, 1) / 1000
                If weights2Object(e, 1) > 100 Then weights2Object(e, 1) = weights2Object(e, 1) / 1000
            Next e
            
            With macroWb.Worksheets("Объекты")
                lastcolumn = .Cells(1, 1).CurrentRegion.Columns.Count
                lastrow = .Cells(1, 1).CurrentRegion.Rows.Count
                Set findCellObject = .Range(.Cells(1, 2), .Cells(lastrow, 2)).Find(currentObject) 'ячейка с текущим объектом на итоговом листе
                If findCellObject Is Nothing Then
                    MsgBox "Не найдено название площадки (определяется по названию файла) из справочника на листе Объекты. Нужно проверить справочник и название файла."
                    GoTo errorExit
                End If
            
                ' Debug.Print findCellObject.Column
                ' Debug.Print findCellObject.Row
                ' Debug.Print macroWb.Worksheets("Объекты").Cells(findCellObject.Row + i, findCellObject.Column + 2).Value

                If minFileDate < 45292 Or maxFileDate > 45657 Then 'проверка ошибок с датами в файле объекта
                    a = MsgBox("В файле " & objectWb.Name & " обнаружены записи не за " & Year(Date) & " год, продолжить?", vbQuestion + vbYesNo + vbDefaultButton2)
                    If a = vbYes Then
                        If minFileDate < 45292 Then minFileDate = CDate("01.01." & Year(Date)) '1 января текущего года
                        If maxFileDate > 45657 Then maxFileDate = CDate("31.12." & Year(Date)) '31 декабря текущего года
                    Else
                        GoTo errorExit
                    End If
                End If
            
                allDates = .Range(.Cells(1, 1), .Cells(1, lastcolumn))
                minDateColumn = 0
                maxDateColumn = 0
                For i = LBound(allDates, 2) To UBound(allDates, 2)
                    If IsDate(allDates(1, i)) Then allDates(1, i) = CDate(allDates(1, i))
                    If allDates(1, i) = minFileDate Then minDateColumn = i - 1
                    If allDates(1, i) = maxFileDate Then maxDateColumn = i - 1
                    If Not minDateColumn = 0 And Not maxDateColumn = 0 Then Exit For
                Next i
                    
                ' Set minDateCell = .Range(.Cells(1, 1), .Cells(1, lastcolumn)).Find(What:=minFileDate, LookIn:=xlValues, LookAt:=xlWhole)
                ' Set maxDateCell = .Range(.Cells(1, 1), .Cells(1, lastcolumn)).Find(What:=maxFileDate, LookIn:=xlValues, LookAt:=xlWhole)
                ' If minDateCell Is Nothing Then Set minDateCell = .Range(.Cells(1, 1), .Cells(1, lastcolumn)).Find(What:=minFileDate, LookIn:=xlFormulas, LookAt:=xlWhole)
                ' If maxDateCell Is Nothing Then Set maxDateCell = .Range(.Cells(1, 1), .Cells(1, lastcolumn)).Find(What:=maxFileDate, LookIn:=xlFormulas, LookAt:=xlWhole)
                
                If Not minDateColumn = 0 And Not maxDateColumn = 0 Then
                    For j = minDateColumn To maxDateColumn + 1 'очистка старых данных по объекту перед заполнением новых
                        For i = 3 To lastrow
                            If .Cells(i, 2) = currentObject Then
                                For n = 0 To landfillsCount - 1
                                    If Sort = True Then .Cells(i + n + 5, j).ClearContents Else .Cells(i + n, j).ClearContents
                                Next n
                            End If
                        Next i
                    Next j

                Else
                    MsgBox "Не найдены минимальная/максимальная даты. minFileDate = " & minFileDate & ", maxFileDate = " & maxFileDate
                    GoTo errorExit
                End If
            
            End With

            Dim sumW1, sumW2 As Double
            sumW1 = 0
            sumW2 = 0
            tempLandfill = Empty
            
            For j = minDateColumn To maxDateColumn + 1 Step 2 'столбец с нужной датой и массой объекта
                For i = 0 To landfillsCount - 1 'цикл по названиям полигонов на итоговом листе в столбце D
                    For e = LBound(weights1Object) To UBound(weights1Object)
                        tempLandfill = Empty
                        For element = 1 To UBound(objects) 'цикл чтобы определить какой полигон написан в реестре объекта
                            If LCase(landfillsOfObject(e, 1)) = LCase(objects(element, 2)) Then tempLandfill = objects(element, 1)
                        Next element
                        If tempLandfill = Empty Then
                            MsgBox "Обнаружено новое название полигона, которого нет в справочнике. Номер строки: " & e + 1 & ". Название: " & landfillsOfObject(e, 1)
                            GoTo errorExit
                        End If
                        If datesOfObject(e, 1) = macroWb.Worksheets("Объекты").Cells(1, j) Then 'нашли столбец с нужной датой
                            If tempLandfill = macroWb.Worksheets("Объекты").Cells(findCellObject.Row + i, findCellObject.Column + 2).Value Then
                                sumW1 = sumW1 + weights1Object(e, 1)
                                sumW2 = sumW2 + weights2Object(e, 1)
                                If Sort = False Then
                                    macroWb.Worksheets("Объекты").Cells(findCellObject.Row + i, j) = sumW1
                                    macroWb.Worksheets("Объекты").Cells(findCellObject.Row + i, j + 1) = sumW2
                                ElseIf Sort = True Then
                                    macroWb.Worksheets("Объекты").Cells(findCellObject.Row + i + landfillsCount, j) = sumW1
                                    macroWb.Worksheets("Объекты").Cells(findCellObject.Row + i + landfillsCount, j + 1) = sumW2
                                End If
                            End If
                        End If
                    Next e
                    sumW1 = 0
                    sumW2 = 0
                Next i
            Next j
            
        End With
    
        objectWb.Close SaveChanges:=False
        fileIndex = fileIndex + 1
    Next 'конец for each
    
    
    With Sheets("Объекты")
        lastRowObj = .Cells(1, 1).CurrentRegion.Rows.Count
        lastColumnObj = .Cells(1, 1).CurrentRegion.Columns.Count
        
        For i = 3 To lastRowObj 'защита от кривых рук
            If Not .Cells(i, 2) = "" Then obj = .Cells(i, 2)
            If Not .Cells(i, 3) = "" Then objType = .Cells(i, 3)
            .Cells(i, lastColumnObj - 1) = obj & .Cells(i, 4) & objType
            .Cells(i, lastColumnObj) = obj
        Next i
        
        yesterdayDate = Sheets("Распределение 1 полугодие").Cells(1, 2).Value - 1
        'yesterdayDate = Date - 1
        
        Dim dates As Variant
        dates = .Range(.Cells(1, 1), .Cells(1, CInt(lastColumnObj))).Value
        For j = LBound(dates) To UBound(dates, 2)
            If dates(1, j) = yesterdayDate Then
                findDateColumn = j
                Exit For
            End If
        Next j
    End With
    
    With Sheets("Распределение 1 полугодие")
        lastRowSplit = .Cells(1, 1).CurrentRegion.Rows.Count
        lastColumnSplit = .Cells(1, 1).CurrentRegion.Columns.Count
        
        Set findLandfillColumnTitle = .Range(.Cells(1, 1), .Cells(lastRowSplit, lastColumnSplit)).Find("Полигон")
        If Not findLandfillColumnTitle Is Nothing Then
            For i = 3 To lastRowSplit 'защита от кривых рук
                If Not .Cells(i, 1) = "" Then obj = .Cells(i, 1)
                If Not .Cells(i, 2) = "" Then objType = .Cells(i, 2)
                .Cells(i, findLandfillColumnTitle.Column + 2) = obj
                .Cells(i, findLandfillColumnTitle.Column + 1) = obj & .Cells(i, findLandfillColumnTitle.Column) & objType
            Next i
        Else
            MsgBox "На листе Распределение 1 полугодие нет столбца с названием Полигон"
            GoTo errorExit
        End If
        
        Set findFactTitle = .Range(.Cells(1, 1), .Cells(lastRowSplit, lastColumnSplit + 20)).Find(What:="Фактический вывоз с объектов", LookIn:=xlValues, LookAt:=xlPart)
        If Not findFactTitle Is Nothing Then
            For i = 3 To lastRowSplit 'как фактически возят
                .Cells(i, findFactTitle.Column) = 0
                For ii = 3 To lastRowObj
                    If .Cells(i, findLandfillColumnTitle.Column + 1) = Sheets("Объекты").Cells(ii, lastColumnObj - 1) Then
                        .Cells(i, findFactTitle.Column) = .Cells(i, findFactTitle.Column) + Sheets("Объекты").Cells(ii, findDateColumn)
                    End If
                Next ii
            Next i
        Else
            MsgBox "На листе Распределение 1 полугодие нет ячейки Фактический вывоз с объектов"
            GoTo errorExit
        End If
    End With
    
    Sheets("Графики").Select
    
    '-------------------- Умные таблицы ---------------------------------------------------------------------------------------
    
    With macroWb.Worksheets("Графики")

        chartTitlesTwoDim = Sheets("Справочник").ListObjects("LandfillsList").ListColumns("Для графиков").DataBodyRange.Value
        ' chartTitles = Array("ВвозНовыйСвет", "ВвозПолигонТБО", "ВвозАвтоБеркут", "ВвозЭкоПлант", "ВвозУКЛО")

        chartTitles = twoDimArrayToOneDim(chartTitlesTwoDim) 'двумерный массив в одномерный
        
        ' Function renamePivotTable(ByVal tbl As Variant, ByVal newName As String, startCounter)
        '     If obj.ShowAutoFilter Then
        '         obj.Name = newName & startCounter
        '         startCounter = startCounter + 1
        '     End If
        ' End Function

        realChartNames = False
        For i = 1 To 2
            counter = 1
            If realChartNames = False Then
                For Each obj In .ListObjects
                    If obj.ShowAutoFilter Then
                        obj.Name = "ВременноеНазвание" & counter
                        counter = counter + 1
                    End If
                    ' renamePivotTable(obj, "ВременноеНазвание" & counter, 1)
                Next obj
            ElseIf realChartNames = True Then
                For Each obj In .ListObjects
                    If obj.ShowAutoFilter Then
                        obj.Name = chartTitles(counter)
                        counter = counter + 1
                    End If
                Next obj
            End If
            realChartNames = True
        Next i

        lastRowChart = .ListObjects(chartTitles(1)).ListRows.Count
        lastDateChart = .ListObjects(chartTitles(1)).ListColumns("Дата").DataBodyRange.Cells(lastRowChart)
        For e = LBound(chartTitles) To UBound(chartTitles)
            tempLastRowChart = .ListObjects(chartTitles(e)).ListRows.Count
            tempLastDateChart = .ListObjects(chartTitles(e)).ListColumns("Дата").DataBodyRange.Cells(tempLastRowChart)
            If Not tempLastRowChart = lastRowChart Then
                MsgBox "В графиках обнаружено разное количество дат. Проверьте все графики, чтобы везде были полностью одинаковые даты."
                GoTo errorExit
            End If
            If Not tempLastDateChart = lastDateChart Then
                MsgBox "В графиках обнаружены несовпащающие последние даты. Проверьте все графики, чтобы везде были полностью одинаковые даты."
                GoTo errorExit
            End If
        Next e

        If lastDateTable > lastDateChart Then
            If lastDateTable = Date And lastDateTable - lastDateChart > 0 Then lastDateTable = lastDateTable - 1
            For i = LBound(chartTitles) To UBound(chartTitles)
                For dayCount = 1 To lastDateTable - lastDateChart
                    .ListObjects(chartTitles(i)).ListRows.Add
                    .ListObjects(chartTitles(i)).ListColumns("Дата").DataBodyRange.Cells(lastRowChart + dayCount).Value = lastDateChart + dayCount
                Next dayCount
            Next i
        End If

        monthTable = Month(CDate(lastDateTable))
        monthChart = Month(CDate(lastDateChart))
        If monthTable > monthChart Then 'если новый месяц то удаляем старые даты для графиков (после добавления новых дат, но текущий месяц на графике определяем заранее)
            For e = LBound(chartTitles) To UBound(chartTitles)
                For n = .ListObjects(chartTitles(e)).ListRows.Count To 1 Step -1
                    tempMonth = Month(CDate(.ListObjects(chartTitles(e)).ListColumns("Дата").DataBodyRange.Cells(n).Value))
                    If tempMonth = monthChart Then .ListObjects(chartTitles(e)).ListRows(n).Delete
                Next n
            Next e
        End If
    '-------------------- Умные таблицы конец ---------------------------------------------------------------------------------------
    
    '-------------------- Графики ---------------------------------------------------------------------------------------
    
        Dim landfillNames As Variant
        landfillNames = Sheets("Справочник").ListObjects("LandfillsList").ListColumns("Полигоны").DataBodyRange.Value
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

        ' For Each chrt In .ChartObjects 'удаление нулевых объектов с легенды
        '     chrt.Legend.Delete
        '     chrt.SetElement(msoElementLegendBottom)
        '     for i = 1 to chrt.Legend.LegendEntries.count step -1
        '         sumObjChart = .ListObjects("titles").ListColumns("Полигон").DataBodyRange.Value
        '         sumObjChart = cdbl(chrt.Name[[#Totals];[chrt.Legend.LegendEntries(i).LegendKey.text]])
        '         ВвозЭкоПлант[[#Totals],[Волхонка АО ""Невский экологический оператор""]]
        '         debug.print i & ": " & sumObjChart
        '         if sumObjChart <= 0 then chrt.Legend.LegendEntries(i).Delete
        '     next i
        ' Next chrt
        
    End With
    
    '-------------------- Графики конец ---------------------------------------------------------------------------------------
    
    
errorExit:
    With Application
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With

End Sub