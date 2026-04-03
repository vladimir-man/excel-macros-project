Option Explicit
' Глобальная константа: относительный путь к папке с накладными
Public Const basePath As String = "Облік/ВХІДНІ НАКЛАДНІ/"

' Шаг 1: формирование словаря ID > накладные
Function GetInvoicesDictionary(wsSource As Worksheet) As Object
    Dim Dict As Object
    Dim LastRow As Long, i As Long
    Dim ID As String, Invoice As String
    
    Set Dict = CreateObject("Scripting.Dictionary")
    LastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    ID = wsSource.Cells(i, 4).Value
    Invoice = wsSource.Cells(i, 6).Value
    
    ' Проверка: если в колонке C уже стоит галочка, пропускаем строку
    If wsSource.Cells(i, 3).Value <> ChrW(&H2713) Then
        If Len(ID) > 0 And Len(Invoice) > 0 Then
            If Not Dict.Exists(ID) Then
                Dict.Add ID, Array(Invoice)
            Else
                Dict(ID) = Join(Dict(ID), ";") & ";" & Invoice
                Dict(ID) = Split(Dict(ID), ";")
                End If
            End If
        End If
    Next i
    
    Set GetInvoicesDictionary = Dict
End Function

' Шаг 2: поиск строки-шапки по ID
Function FindIDRow(ws As Worksheet, ByVal ID As String) As Long
    Dim rng As Range
    Set rng = ws.Columns(1).Find(What:=ID, LookAt:=xlWhole, LookIn:=xlValues)
    
    If Not rng Is Nothing Then
        FindIDRow = rng.Row
    Else
        FindIDRow = 0
    End If
End Function

' Шаг 3: проверка/создание/обновление подшапки


Sub CheckOrCreateSubRows(ByVal IDRow As Long, ByVal Invoices As Variant)
    Dim ws As Worksheet
    Dim StartRow As Long, writeRow As Long
    Dim inv As Variant
    Dim emptyCount As Long, neededCount As Long, addCount As Long
    
    Set ws = ThisWorkbook.Sheets("Тренировка")
    StartRow = IDRow + 1
    
    ' Защита от пустого массива
    If IsEmpty(Invoices) Then Exit Sub
    
    ' --- Шаг 1: обновляем шапку через конкатенацию ---
    For Each inv In Invoices
        If Len(ws.Cells(IDRow, 2).Value) = 0 Then
            ws.Cells(IDRow, 2).Value = inv
        Else
            ws.Cells(IDRow, 2).Value = ws.Cells(IDRow, 2).Value & "; " & inv
        End If
    Next inv
    
    ' --- Шаг 2: проверяем наличие группы ---
    If ws.Rows(StartRow).OutlineLevel > ws.Rows(IDRow).OutlineLevel Then
        emptyCount = 0
        writeRow = StartRow
        Do While ws.Rows(writeRow).OutlineLevel > ws.Rows(IDRow).OutlineLevel
            If Len(Trim(ws.Cells(writeRow, 2).Value)) = 0 Then
                emptyCount = emptyCount + 1
            End If
            writeRow = writeRow + 1
        Loop
        
        neededCount = UBound(Invoices) + 1
        addCount = neededCount - (emptyCount - 1)
        
        If addCount > 0 Then
            ws.Rows(writeRow & ":" & writeRow + addCount - 1).Insert Shift:=xlDown
        End If
        
        ' ?? Подшапка: запись как гиперссылки
        writeRow = StartRow
        For Each inv In Invoices
            Do While Len(Trim(ws.Cells(writeRow, 2).Value)) <> 0
                writeRow = writeRow + 1
            Loop
            ws.Cells(writeRow, 2).Formula = _
                "=HYPERLINK(""" & basePath & inv & """,""" & inv & """)"
        Next inv
        
    Else
        neededCount = UBound(Invoices) + 1
        ws.Rows(StartRow & ":" & StartRow + neededCount).Insert Shift:=xlDown
        
        writeRow = StartRow
        For Each inv In Invoices
            ' ?? Подшапка: запись как гиперссылки
            ws.Cells(writeRow, 2).Formula = _
                "=HYPERLINK(""" & basePath & inv & """,""" & inv & """)"
            writeRow = writeRow + 1
        Next inv
        
        ws.Cells(writeRow, 2).Value = "" ' пустая строка
        ws.Rows(StartRow & ":" & writeRow).Group
        
        If ws.Rows(StartRow).OutlineLevel > ws.Rows(IDRow).OutlineLevel Then
            On Error Resume Next
            ws.Rows(StartRow).ShowDetail = True
            On Error GoTo 0
        End If
    End If
End Sub

' Главный макрос: связывает шаги 1–3
Sub RecordOrdersNumberInTable()
    Dim wsTarget As Worksheet
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim Dict As Object
    Dim Key As Variant
    Dim IDRow As Long
    Dim FileName As Variant
    Dim rng As Range
    
    ' Выбор файла через диалог
    FileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")
    If FileName = False Then Exit Sub
    
    ' Открываем выбранный файл
    Set wbSource = Application.Workbooks.Open(FileName)
    Set wsSource = wbSource.Sheets("РЕЕСТР вх накл")
    
    ' Основной лист в текущей книге
    Set wsTarget = ThisWorkbook.Sheets("Тренировка")
    
    ' Формируем словарь из выбранного файла
    Set Dict = GetInvoicesDictionary(wsSource)
    
    ' Перебираем все ID
    For Each Key In Dict.Keys
        IDRow = FindIDRow(wsTarget, Key)
        If IDRow > 0 Then
            Call CheckOrCreateSubRows(IDRow, Dict(Key))
            
            ' Ставим галочку в реестре
          Dim firstAddress As String
            Set rng = wsSource.Columns(4).Find(What:=Key, LookAt:=xlWhole, LookIn:=xlValues)

            If Not rng Is Nothing Then
                firstAddress = rng.Address
                Do
                    With wsSource.Cells(rng.Row, 3)   ' галочка в колонке C
                        .Value = ChrW(&H2713)
                        .Font.Bold = True
                        .Font.Color = vbGreen
                    End With
                    Set rng = wsSource.Columns(4).FindNext(rng)
                Loop While Not rng Is Nothing And rng.Address <> firstAddress
            End If
        End If
    Next Key
    
    ' Сохраняем изменения в основной книге
    wsTarget.Parent.Save
    
    ' Закрываем выбранный файл с сохранением галочек
    wbSource.Close SaveChanges:=True
End Sub

