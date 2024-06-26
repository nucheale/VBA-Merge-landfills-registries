Public NS#

Function twoDimArrayToOneDim(oldArr) 'двумерный массив в одномерный
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For I = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(I) = oldArr(I, 1)
    Next I
    twoDimArrayToOneDim = newArr
End Function

Function removeDublicatesFromOneDimArr(arr) 'удаление дубликатов в одномерном массиве
    Dim coll As New Collection
    For Each e In arr
        On Error Resume Next
        coll.Add e, e
        On Error GoTo 0
    Next e
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To coll.Count)
    For I = 1 To coll.Count
        uniqueArr(I) = coll(I)
    Next I
    removeDublicatesFromOneDimArr = uniqueArr
End Function

Private Sub CommandButton1_Click()
    
    If Len(Left(TextBox1.Value, 1) & Left(ComboBox1.Value, 1) & Left(TextBox3.Value, 1)) < 3 Then
        MsgBox "Заполнены не все поля"
        GoTo errorExit
    End If
    
    If TextBox3.Value <= 0 Then GoTo errorExit2
    
    On Error GoTo errorExit2
    entryDate = CDate(TextBox1.Value)
    entryLandfill = CStr(ComboBox1.Value)
    entryWeight = CDbl(Replace(TextBox3.Value, ".", ",")) * 1000 ' т в кг
    On Error GoTo 0

    With Sheets("Вывоз")
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        .Cells(lastRow + 1, 1) = entryDate
        .Cells(lastRow + 1, 8) = entryLandfill
        .Cells(lastRow + 1, 9) = entryWeight
        .Cells(lastRow + 1, 10) = entryWeight
    End With

    ComboBox1.Value = Empty
    TextBox3.Value = Empty
errorExit:
Exit Sub

errorExit2:
MsgBox "Введено неверное значение"
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton3_Click()
    UserForm2_Calendar.Show
End Sub

Private Sub ComboBox1_Change()
    If ComboBox1.Value = "Новый Свет" Then TextBox3.Value = NS Else TextBox3.Value = Empty
End Sub


Private Sub UserForm_Initialize()
    NS = 723.287671232877
    TextBox1.Value = Format(Date - 1, "dd.mm.yyyy")
    TextBox3.Value = Empty
    With ComboBox1
        ' With Sheets("Вывоз")
        '     lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        '     landfills = .Range(.Cells(2, 8), .Cells(lastRow, 8))
        '     landfills = twoDimArrayToOneDim(landfills)
        '     landfills = removeDublicatesFromOneDimArr(landfills)
        ' End With
        landfills = Array("Новый Свет", "Эко Плант", "Полигон ТБО", "Авто-Беркут", "УКЛО")
        If Not .ListCount = UBound(landfills) + 1 Then
            For Each e In landfills
                .AddItem e
            Next e
        End If
    End With
End Sub
