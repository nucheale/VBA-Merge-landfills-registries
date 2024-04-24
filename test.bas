'contacts: i.zabolotny@spb-neo.ru, telegram @kxcvt

Sub STATS()
       
    With Application
        .ScreenUpdating = False 
        .Calculation = xlCalculationManual
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With

    With Sheets("Входящие")
    ' Sheets("Входящие").Select
        If Not .Cells(1, 10) Like "[Пп]ричина обращения" Then
            MsgBox "Проверьте корректность листа Входящие"
            Goto errorExit
        End If

    lastRowOrders = .Cells(Rows.Count, 4).End(xlUp).Row
    For i = lastRowOrders To 1 Step -1
        If .Cells(i, 4) = "" Then 'убираем строки с пустым адресом и пустой причиной обращения
            If .Cells(i, 10) = "" Then
                .Rows(i).EntireRow.Delete
            Else
                MsgBox ("Не во всех строках заполнен адрес, при этом указана причина обращения. Заполните и попробуйте еще раз.")
                Goto errorExit
            End If
        End If
        If Cells(i, 10) = "" Then 'проверяем заполнение причин обращения
            MsgBox ("Не во всех строках заполнена причина обращения. Заполните и попробуйте еще раз.")
            Goto errorExit
        End If
    Next i
    
    
    With Sheets(1)
    Set columnName = .Rows(2).Find("Динамика к предыдущему дню")
    lastColumnIndex = CInt(columnName.Column)
    If Not columnName Is Nothing Then
        .Columns(lastColumnIndex).Insert
        .Cells(2, lastColumnIndex).Value = CLng(Date)
        .Columns(lastColumnIndex - 5).Hidden = True
    Else
        MsgBox "Не найден столбец с названием 'Динамика к предыдущему дню'"
        Goto errorExit
    End If
    
    lastDayIndex = WorksheetFunction.CountA(Rows(2)) - 1 'вообще можно использовать lastColumnIndex без изменения, но это чтобы не запутаться
    
    For i = 61 To 63
        Cells(i, lastDayIndex).FormulaR1C1 = Cells(i, lastDayIndex - 1).FormulaR1C1
        Cells(i, lastDayIndex).Borders.LineStyle = True
    Next i
    
    For i = 65 To 68
        Cells(i, lastDayIndex).FormulaR1C1 = Cells(i, lastDayIndex - 1).FormulaR1C1
        Cells(i, lastDayIndex).Borders.LineStyle = True
    Next i
    
                
    Sheets(1).Select 'заполняем счетесли
    For i = 5 To 59 Step 3
        ActiveSheet.Cells(i, lastDayIndex).FormulaR1C1 = "=COUNTIFS(Входящие!C2,Статистика!RC3,Входящие!C10,""*жалоба*"")+COUNTIFS(Входящие!C2,Статистика!RC3,Входящие!C10,""*нет контейнер*"")+COUNTIFS(Входящие!C2,Статистика!RC3,Входящие!C10,""*вывезли не все*"")"
    Next i
    For i = 4 To 58 Step 3
        ActiveSheet.Cells(i, lastDayIndex).FormulaR1C1 = "=COUNTIFS(Входящие!C2,Статистика!RC3,Входящие!C3,""Юр. лицо"",Входящие!C10,""*жалоба*"")+COUNTIFS(Входящие!C2,Статистика!RC3,Входящие!C3,""Юр. лицо"",Входящие!C10,""*нет контейнер*"")+COUNTIFS(Входящие!C2,Статистика!RC3,Входящие!C3,""Юр. лицо"",Входящие!C10,""*вывезли не все*"")"
    Next i
    For i = 3 To 57 Step 3
        ActiveSheet.Cells(i, lastDayIndex).FormulaR1C1 = "=R[2]C-R[1]C"
    Next i
    
    Range(Cells(3, lastDayIndex), Cells(59, lastDayIndex)).Value = Range(Cells(3, lastDayIndex), Cells(59, lastDayIndex)).Value 'вставляем данные как значения

    Set columnName = Sheets(1).Rows(2).Find("Динамика к предыдущему дню")
    lastColumnIndex = CInt(columnName.Column)
    
    On Error Resume Next 'заполняем проценты, игнорируя ошибку деления на ноль
    For i = 3 To 59
        Sheets(1).Cells(i, lastColumnIndex) = (0 - (100 - ActiveSheet.Cells(i, lastDayIndex) * 100 / ActiveSheet.Cells(i, lastDayIndex - 1))) / 100
        If Sheets(1).Cells(i, lastDayIndex - 1) = 0 Then
            If Sheets(1).Cells(i, lastDayIndex) = 0 Then
                Sheets(1).Cells(i, lastColumnIndex) = 0
            Else
                Sheets(1).Cells(i, lastColumnIndex) = 1
            End If
        End If
    Next i
    
    Cells(62, lastColumnIndex) = (0 - (100 - ActiveSheet.Cells(63, lastDayIndex) * 100 / ActiveSheet.Cells(63, lastDayIndex - 1))) / 100
    If Cells(62, lastDayIndex - 1) = 0 Then
        If Cells(62, lastDayIndex) > 0 Then
            Cells(62, lastColumnIndex) = 1
        Else
            Cells(62, lastColumnIndex) = 0
        End If
    End If
    
    For i = 65 To 68 'проценты по генподрядчикам
    Cells(i, lastColumnIndex) = (0 - (100 - ActiveSheet.Cells(i, lastDayIndex) * 100 / ActiveSheet.Cells(i, lastDayIndex - 1))) / 100
     If Cells(i, lastDayIndex - 1) = 0 Then
        If Cells(i, lastDayIndex) = 0 Then
         Cells(i, lastColumnIndex) = 0
        Else
         Cells(i, lastColumnIndex) = 1
        End If
     End If
    Next i
    
    Cells(61, lastColumnIndex).Interior.Color = Cells(62, lastColumnIndex).DisplayFormat.Interior.Color 'заливка
    Cells(63, lastColumnIndex).Interior.Color = Cells(62, lastColumnIndex).DisplayFormat.Interior.Color
    
    On Error GoTo 0


    Rows("1:63").Hidden = False 'показываем все строки
    
    For i = 5 To 59 Step 3
        Cells(i, 1) = Cells(i, lastDayIndex)
    Next i
    For i = 4 To 58 Step 3
        Cells(i, 1) = Cells(i + 1, lastDayIndex)
    Next i
    For i = 3 To 57 Step 3
        Cells(i, 1) = Cells(i + 2, lastDayIndex)
    Next i
    
    With ActiveSheet.Sort
        .SortFields.clear
        .SortFields.Add Key:=Range("A3:A59")
        .SortFields.Add Key:=Range("B3:B59")
        .Header = xlYes
        .SetRange Range(Cells(2, 1), Cells(59, lastColumnIndex))
        .Apply
    End With
    
    
    For i = 3 To 59 Step 3 'скрываем лишние строки и столбцы
        Rows(i).Hidden = True
    Next i
    For i = 4 To 58 Step 3
        Rows(i).Hidden = True
    Next i


    applicationsSchedule = WorksheetFunction.CountIf(Sheets("Входящие").Range("J1:J1000"), "*изменение графика*") 'считаем сумму обращений по изменению графика и отмене вывоза
    applicationsCancel = WorksheetFunction.CountIf(Sheets("Входящие").Range("J1:J1000"), "*отмена вывоза*")
    applicationsOrder = (WorksheetFunction.CountIf(Sheets("Входящие").Range("J1:J1000"), "*заявка на*")) + (WorksheetFunction.CountIf(Sheets("Входящие").Range("J1:J1000"), "*замена контейнер*"))
    applicationsNewKp = WorksheetFunction.CountIf(Sheets("Входящие").Range("J1:J1000"), "*Новая КП, добавить*")
    applicationsKpChange = WorksheetFunction.CountIf(Sheets("Входящие").Range("J1:J1000"), "*Изменения в контейнерной группе*")
    
    Sheets(1).Cells(70, 3) = "Обращений по изменению графика: " & applicationsSchedule
    Sheets(1).Cells(71, 3) = "Обращений по отмене вывоза: " & applicationsCancel
    Sheets(1).Cells(72, 3) = "Заявок на вывоз: " & applicationsOrder
    Sheets(1).Cells(73, 3) = "Новых КП: " & applicationsNewKp
    'Sheets(1).Cells(74, 3) = "Изменений в КП: " & kpChange
    
    
    Range(Cells(70, 3), Cells(73, lastColumnIndex - 5)).Borders(xlEdgeBottom).Weight = xlThin 'внешние границы ячеек
    Range(Cells(70, 3), Cells(73, lastColumnIndex - 5)).Borders(xlEdgeTop).Weight = xlThin
    Range(Cells(70, 3), Cells(73, lastColumnIndex - 5)).Borders(xlEdgeLeft).Weight = xlThin
    Range(Cells(70, 3), Cells(73, lastColumnIndex - 5)).Borders(xlEdgeRight).Weight = xlThin
    
    
    Sheets("fullstats").Select 'лист fullstats
    lastColumnIndexFS = Cells(3, Columns.Count).End(xlToLeft).Column
    n = 1
    For j = lastColumnIndexFS + 1 To lastColumnIndexFS + 5 'порядковые номера столбцов и дата
        Cells(1, j).Value = Cells(1, lastColumnIndexFS) + n
        Cells(2, j).Value = CLng(Date)
        n = n + 1
    Next j
    Cells(3, lastColumnIndexFS + 1).Value = "Жалоба" 'название столбца
    Cells(3, lastColumnIndexFS + 2).Value = "Заявка" 'название столбца
    Cells(3, lastColumnIndexFS + 3).Value = "График" 'название столбца
    Cells(3, lastColumnIndexFS + 4).Value = "Отмена" 'название столбца
    Cells(3, lastColumnIndexFS + 5).Value = "Новая КП" 'название столбца

        
    For i = 6 To 60 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 1).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*жалоба*"")+COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*нет контейнер*"")+COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*вывезли не все*"")" 'итог жалобы по районам
    Next i
    For i = 5 To 59 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 1).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*жалоба*"")+COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*нет контейнер*"")+COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*вывезли не все*"")" 'итог жалобы юр. лиц по районам
    Next i
    For i = 4 To 58 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 1).FormulaR1C1 = "=R[2]C-R[1]C" 'итог жалобы жил. фонда по районам
    Next i
    For i = 6 To 60 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 2).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*заявка на*"")+COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*замена контейнер*"")" 'итог заявки по районам
    Next i
    For i = 5 To 59 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 2).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*заявка на*"")+COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*замена контейнер*"")" 'итог заявки юр. лиц по районам
    Next i
    For i = 4 To 58 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 2).FormulaR1C1 = "=R[2]C-R[1]C" 'итог заявки жил. фонда по районам
    Next i
    For i = 6 To 60 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 3).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*изменение графика*"")" 'итог график по районам
    Next i
    For i = 5 To 59 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 3).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*изменение графика*"")" 'итог график юр. лиц по районам
    Next i
    For i = 4 To 58 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 3).FormulaR1C1 = "=R[2]C-R[1]C" 'итог график жил. фонда по районам
    Next i
    For i = 6 To 60 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 4).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*отмена вывоза*"")" 'итог отмены по районам
    Next i
    For i = 5 To 59 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 4).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*отмена вывоза*"")" 'итог отмены юр. лиц по районам
    Next i
    For i = 4 To 58 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 4).FormulaR1C1 = "=R[2]C-R[1]C" 'итог отмены жил. фонда по районам
    Next i
    For i = 6 To 60 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 5).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C10,""*новая КП, добавить*"")" 'итог новые по районам
    Next i
    For i = 5 To 59 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 5).FormulaR1C1 = "=COUNTIFS(Входящие!C2,fullstats!RC1,Входящие!C3,""Юр. лицо"",Входящие!C10,""*новая КП, добавить*"")" 'итог новые юр. лиц по районам
    Next i
    For i = 4 To 58 Step 3
        ActiveSheet.Cells(i, lastColumnIndexFS + 5).FormulaR1C1 = "=R[2]C-R[1]C" 'итог новые жил. фонда по районам
    Next i
    
    Range(Cells(4, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Value = Range(Cells(4, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Value 'вставляем данные как значения
        
    Range(Cells(1, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Borders.LineStyle = xlContinuous 'границы одного дня
    Range(Cells(1, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Borders(xlEdgeBottom).Weight = xlMedium 'жирные границы одного дня
    Range(Cells(1, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Borders(xlEdgeTop).Weight = xlMedium
    Range(Cells(1, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Borders(xlEdgeLeft).Weight = xlMedium
    Range(Cells(1, lastColumnIndexFS + 1), Cells(60, lastColumnIndexFS + 5)).Borders(xlEdgeRight).Weight = xlMedium

    
    If Sheets(1).Cells(63, lastDayIndex) = 0 Then 'проверка если 0 жалоб то все зеленое
        Sheets("allGreen").Select
        Cells(1, 2) = "Количество обращений на горячую линию регоператора по невывозу ТКО"
        For i = 2 To 68
            Sheets("allGreen").Cells(i, 2) = Sheets(1).Cells(i, 3)
            Sheets("allGreen").Cells(i, 3) = Sheets(1).Cells(i, 4)
            Sheets("allGreen").Cells(i, 4) = Sheets(1).Cells(i, lastDayIndex - 4)
            Sheets("allGreen").Cells(i, 5) = Sheets(1).Cells(i, lastDayIndex - 3)
            Sheets("allGreen").Cells(i, 6) = Sheets(1).Cells(i, lastDayIndex - 2)
            Sheets("allGreen").Cells(i, 7) = Sheets(1).Cells(i, lastDayIndex - 1)
            Sheets("allGreen").Cells(i, 8) = Sheets(1).Cells(i, lastDayIndex)
            Sheets("allGreen").Cells(i, 9) = Sheets(1).Cells(i, lastColumnIndex)
        Next i
    
        For i = 70 To 73
            Sheets("allGreen").Cells(i, 2) = Sheets(1).Cells(i, 3)
        Next i
        
        lastColumnIndexAllGreen = WorksheetFunction.CountA(Rows(2))
        Sheets("allGreen").Range(Cells(1, 2), Cells(73, lastColumnIndexAllGreen + 1)).ExportAsFixedFormat Type:=xlTypePDF, Filename:="Y:\Отдел взаимодействия с перевозчиками\Статистика по обращениям\Статистика " & Format(Now, "DD.MM.YYYY") & ".pdf", OpenAfterPublish:=True
        Sheets("allGreen").Range("B:Z").ClearContents
        Sheets(1).Select
        
    Else
        
        Sheets(1).Select
        pdfFileName = ActiveWorkbook.Path & "\" & "Статистика " & Format(Now, "DD.MM.YYYY") & ".pdf"
        Sheets(1).Range(Cells(1, 3), Cells(73, lastColumnIndex)).ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, OpenAfterPublish:=True
        'создаем PDF
        
    End If
    
    
    
    newFileName = ActiveWorkbook.Path & "\" & "Статистика " & Format((Now() + 1), "DD.MM.YYYY") & ".xlsm"
    ActiveWorkbook.SaveAs Filename:=newFileName

errorExit:
    With Application
        .ScreenUpdating = True 
        .Calculation = xlCalculationAutomatic
        .AskToUpdateLinks = True
        .DisplayAlerts = True
    End With

End Sub


