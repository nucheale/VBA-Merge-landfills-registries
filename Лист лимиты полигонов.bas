Sub Пересчет_листа_лимит_полигонов()

    Application.Calculation = xlCalculationManual
    
    With Sheets("Объекты")
        lastRowObjects = .Cells(1, 1).CurrentRegion.Rows.Count
        lastColumnObjects = .Cells(1, 1).CurrentRegion.Columns.Count
        Dim objectsData As Variant
        objectsData = .Range(.Cells(1, 1), .Cells(lastRowObjects, lastColumnObjects))
    End With
    
    With Sheets("Лимиты полигонов")

        lastRow1 = .Cells(1, 1).CurrentRegion.Rows.Count
        lastColumn1 = .Cells(1, 1).CurrentRegion.Columns.Count
    
        lastRow2 = .Cells(lastRow1 + 1, 1).CurrentRegion.Rows.Count
        lastColumn2 = .Cells(lastRow1 + 1, 1).CurrentRegion.Columns.Count
    
        monthHalf1 = Array("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь")
        monthHalf2 = Array("Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
         
        Set cellLandfills = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Объект захоронения")
        Set cellAverageImport = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Среднесуточный ввоз")
        Set cellLimitHalf1 = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Годовой лимит 1 полугодие")
        Set cellLimitHalf2 = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Годовой лимит 2 полугодие")
        Set cellLimitResidueHalf1 = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Остаток лимита 1 полугодие")
        Set cellLimitResidueHalf2 = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Остаток лимита 2 полугодие")
        Set cellImportHalf1 = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Ввезено 1 полугодие")
        Set cellImportHalf2 = .Range(.Cells(1, 1), .Cells(lastRow1, lastColumn1)).Find("Ввезено 2 полугодие")
    
        titles1 = Array(cellLandfills, cellAverageImport, cellLimitHalf1, cellLimitHalf2, cellLimitResidueHalf1, cellLimitResidueHalf2, cellImportHalf1, cellImportHalf2)
    
        For i = LBound(titles1) To UBound(titles1)
            If titles1(i) Is Nothing Then GoTo errorExit
        Next i
        
        Dim landfillNames As Variant
        landfillNames = Sheets("Справочник").ListObjects("LandfillsList").ListColumns("Полигоны").DataBodyRange.Value
    
        For i = 1 To UBound(landfillNames, 1)
            .Cells(cellLandfills.Row + i, cellLandfills.Column) = landfillNames(i, 1)
        Next i
    
    
        Set cellDate = .Range(.Cells(lastRow1 + 1, 1), .Cells(lastRow2, lastColumn2)).Find("Дата")
        
        For j = 1 To UBound(landfillNames, 1)
            .Cells(cellDate.Row, cellDate.Column + j) = landfillNames(j, 1)
        Next j

    
    
        Dim landfillsSumData As Variant
        landfillsSumData = .Range(.Cells(cellDate.Row + 1, cellDate.Column), .Cells(lastRow2, lastColumn2))
        firstMonthDayRow = 1
        For i = 1 To UBound(landfillsSumData, 1)
            If IsDate(landfillsSumData(i, 1)) Then
                For j = 1 To UBound(landfillNames, 1)
                    For e = 1 To UBound(objectsData, 2)
                        landfillDateSum = 0
                        If landfillsSumData(i, 1) = objectsData(1, e) And objectsData(2, e) = "Масса полигона" Then
                            For ii = 1 To UBound(objectsData, 1)
                                If objectsData(ii, 4) = landfillNames(j, 1) Then
                                    landfillDateSum = landfillDateSum + objectsData(ii, e)
                                    landfillsSumData(i, j + 1) = landfillDateSum
                                End If
                            Next ii
                        End If
                    Next e
                Next j
            Else
                For j = 1 To UBound(landfillNames, 1)
                    landfillSum = 0
                    For ii = firstMonthDayRow To i - 1
                        landfillSum = landfillSum + landfillsSumData(ii, j + 1)
                        landfillsSumData(i, j + 1) = landfillSum
                    Next ii
                Next j
            firstMonthDayRow = i + 1
            End If
        Next i
    
        .Cells(cellDate.Row + 1, cellDate.Column).Resize(UBound(landfillsSumData, 1), UBound(landfillsSumData, 2)) = landfillsSumData
    
    End With
    
    With Application
   '     .Calculation = xlCalculationAutomatic
    End With

    Exit Sub

errorExit:
    With Application
        .Calculation = xlCalculationAutomatic
    End With
    MsgBox "Ошибка"

End Sub