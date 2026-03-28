Attribute VB_Name = "Module9"
Sub RecordOrdersNumberInTable()
    Dim fd As FileDialog
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim id As Variant, invoice As String
    Dim targetRow As Variant
    
    ' 1. Выбор исходного файла
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Выберите ИСХОДНЫЙ файл Excel"
        .Filters.Add "Файлы Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb", 1
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    
    ' 2. Открытие исходного файла
    Set wbSource = Workbooks.Open(filePath)
    Set wsSource = wbSource.Worksheets("РЕЕСТР вх накл")
    
    ' 3. Целевой лист
    Set wsTarget = ThisWorkbook.Sheets("Тренировка")
    
    ' 4. Определяем последнюю строку в реестре
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    
    ' 5. Сбор и запись данных
    For i = 2 To lastRow
        id = wsSource.Cells(i, "D").Value
        invoice = wsSource.Cells(i, "F").Value
        
        ' Проверяем: есть ли ID и накладная, и нет отметки
        If id <> "" And id <> 0 And invoice <> "" And wsSource.Cells(i, "C").Value = "" Then

            
            ' Ищем ID в целевом листе
            On Error Resume Next
            targetRow = Application.Match(id, wsTarget.Columns(1), 0)
            On Error GoTo 0
            
            If Not IsError(targetRow) Then
                ' ID найден > дописываем накладную
                If wsTarget.Cells(targetRow, 2).Value <> "" Then
                    wsTarget.Cells(targetRow, 2).Value = wsTarget.Cells(targetRow, 2).Value & ", " & invoice
                Else
                    wsTarget.Cells(targetRow, 2).Value = invoice
                End If
                
                ' Ставим галочку
            With wsSource.Cells(i, 3)
                .Value = ChrW(&H2713)
                .Font.Bold = True
                .Font.Color = vbGreen
            End With

            End If
        End If
    Next i
    
    MsgBox "Обработка завершена."
End Sub
