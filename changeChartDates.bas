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
    End With

errorExit:
    With Application
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

