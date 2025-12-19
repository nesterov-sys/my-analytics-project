Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateAllReportsSideBySide()
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        MsgBox "Нет активной рабочей книги!", vbCritical
        Exit Sub
    End If

    Dim wsTD As Worksheet, wsReport As Worksheet
    Dim lastRowTD As Long, lastDataRow As Long
    Dim i As Long, j As Long, colOffset As Long, reportType As Long
    Dim metrics As Variant, headers() As String
    Dim DATA_START_ROW As Long: DATA_START_ROW = 11
    Dim hoursValue As Double
    Dim col As Long

    On Error Resume Next
    hoursValue = Application.InputBox("Введите количество часов для расчёта:", "Часы работы", 11, Type:=1)
    If hoursValue = 0 Then Exit Sub
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If Not SheetExists("TDSheet", ActiveWorkbook) Then
        MsgBox "Лист TDSheet не найден в активной книге!", vbExclamation
        GoTo CleanUp
    End If

    Set wsReport = CreateOrClearReportSheet("Сводный отчет", ActiveWorkbook)
    Set wsTD = ActiveWorkbook.Sheets("TDSheet")

    lastRowTD = wsTD.Cells(wsTD.Rows.Count, "A").End(xlUp).Row
    For i = DATA_START_ROW To lastRowTD
        If wsTD.Cells(i, 1).Value = "Итого" Then
            lastRowTD = i - 1
            Exit For
        End If
    Next i

    If lastRowTD < DATA_START_ROW Then
        MsgBox "Нет данных для обработки!", vbExclamation
        GoTo CleanUp
    End If

    ' ===МЕТРИКИ И НОРМЫ ===
    ReDim metrics(0 To 24)
    metrics(0) = "01 Отбор ECOM (шт)"
    metrics(1) = "02 Групповой отбор ECOM (шт)"
    metrics(2) = "03 Отбор ECOM экспресс (шт)"
    metrics(3) = "04 Отбор контейнера (конт)"
    metrics(4) = "05 Отбор ТТ ( магазин+ТЕС+Реклама) (шт)"
    metrics(5) = "06 Перемещение товара (конт)"
    metrics(6) = "07 Перемещение контейнеров (конт)"
    metrics(7) = "08 Приемка без КИЗ (шт)"
    metrics(8) = "09 Приемка с КИЗ (шт)"
    metrics(9) = "10_1 Размещение (шт)"
    metrics(10) = "11 Размещение (конт)"
    metrics(11) = "12_1 Приемка КД на ПК (шт)"
    metrics(12) = "12_2 Приемка КД на ТСД (шт)"
    metrics(13) = "13 Приемка КД (конт)"
    metrics(14) = "14 Проклейка (шт)"
    metrics(15) = "15 Расформирование ECOM (шт)"
    metrics(16) = "16 Упаковка ТТ (шт)"
    metrics(17) = "17 Сортировка ECOM (шт)"
    metrics(18) = "18 Упаковка ECOM внешняя КС (заказ)"
    metrics(19) = "19 Упаковка ECOM своя КС (заказ)"
    metrics(20) = "20 Стикеровка (шт)"
    metrics(21) = "21 Инвентаризация (шт)"
    metrics(22) = "22 Доп. работы (мин)"
    metrics(23) = "23 Копакинг (набор шт.)"
    metrics(24) = "24 Коробочное размещение (кор)"

    ' Нормы — metrics
    Dim norms As Variant
    ReDim norms(0 To 24)
    norms(0) = 80
    norms(1) = 90
    norms(2) = 57
    norms(3) = 120
    norms(4) = 300
    norms(5) = 25
    norms(6) = 110
    norms(7) = 400
    norms(8) = 120
    norms(9) = 400
    norms(10) = 100
    norms(11) = 200
    norms(12) = 0
    norms(13) = 80
    norms(14) = 240
    norms(15) = 50
    norms(16) = 400
    norms(17) = 150
    norms(18) = 150
    norms(19) = 200
    norms(20) = 0
    norms(21) = 650
    norms(22) = 60
    norms(23) = 350
    norms(24) = 60

    ' === ЗАГОЛОВКИ: 53 элемента (0–52) ===
    ReDim headers(0 To 52)
    headers(0) = "Сотрудник"
    headers(1) = "Часы"

    ' Метрики: 25 шт > headers(2) to headers(26)
    For j = 0 To 24
        headers(j + 2) = metrics(j)
    Next j

    ' Нормы: 25 шт > headers(27) to headers(51)
    For j = 0 To 24
        headers(j + 27) = "Норма " & (j + 1)
    Next j

    ' Итог > headers(52)
    headers(52) = "Итог"

    For reportType = 1 To 3
        Select Case reportType
            Case 1: colOffset = 0   ' Таблица 1
            Case 2: colOffset = 56  ' Таблица 2
            Case 3: colOffset = 112 ' Таблица 3
        End Select

        wsReport.Cells(1, 2 + colOffset).Value = IIf(reportType = 1, "Штат", IIf(reportType = 2, "Аутсорс", "Все сотрудники"))
        wsReport.Cells(1, 2 + colOffset).Font.Bold = True
        wsReport.Cells(1, 2 + colOffset).Font.Size = 12

        ' === ВЫВОД ЗАГОЛОВКОВ ===
        wsReport.Cells(3, 1 + colOffset).Value = "№"
        With wsReport.Columns(1 + colOffset)
            .ColumnWidth = 5
            .HorizontalAlignment = xlCenter
        End With

        For j = 0 To 52
            wsReport.Cells(3, j + 2 + colOffset).Value = headers(j)
        Next j

        With wsReport.Range(wsReport.Cells(3, 1 + colOffset), wsReport.Cells(3, 54 + colOffset))
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
            .HorizontalAlignment = xlCenter
        End With

        Dim outputRow As Long: outputRow = 4

        For i = DATA_START_ROW To lastRowTD
            Dim empName As String, hasData As Boolean
            empName = wsTD.Cells(i, 1).Value
            hasData = (wsTD.Cells(i, 4).Value <> 0 And wsTD.Cells(i, 4).Value <> "")

            Select Case reportType
                Case 1: If Not hasData Then GoTo NextEmployee
                Case 2: If hasData Then GoTo NextEmployee
            End Select

            wsReport.Cells(outputRow, 2 + colOffset).Value = empName
            wsReport.Cells(outputRow, 3 + colOffset).Value = hoursValue

            ' Нормы в столбце 29 > 29 + colOffset
            Dim normStartCol As Long
            normStartCol = 29 + colOffset

            ' нормы (столбцы 29–53)
            For j = 0 To 24
                wsReport.Cells(outputRow, j + 29 + colOffset).Value = hoursValue * norms(j)
            Next j

            wsTD.Rows(8).UnMerge

            ' формулы выполнения (%)
            For j = 0 To 24
                Dim foundCell As Range
                Set foundCell = wsTD.Range("A8:XFD8").Find(metrics(j), LookIn:=xlValues, LookAt:=xlWhole)

                If Not foundCell Is Nothing Then
                    Dim targetColumn As Long
                    Dim offsetRow As Long: offsetRow = 0

                    Select Case metrics(j)
                        Case "04 Отбор контейнера (конт)", _
                             "06 Перемещение товара (конт)", _
                             "07 Перемещение контейнеров (конт)", _
                             "11 Размещение (конт)", _
                             "13 Приемка КД (конт)", _
                             "23 Копакинг (набор шт.)", _
                             "24 Коробочное размещение (кор)"
                            targetColumn = foundCell.Column + 1
                        Case Else
                            targetColumn = foundCell.Column
                    End Select

                    Dim targetCell As Range
                    Set targetCell = wsTD.Cells(i + offsetRow, targetColumn)

                    Dim formulaStr As String
                    formulaStr = "=IF(" & wsReport.Cells(outputRow, j + 29 + colOffset).Address(False, False) & "<>0," & _
                                 "'" & wsTD.Name & "'!" & targetCell.Address(False, False) & "/" & _
                                 wsReport.Cells(outputRow, j + 29 + colOffset).Address(False, False) & ",0)"
                    wsReport.Cells(outputRow, j + 4 + colOffset).Formula = formulaStr
                Else
                    wsReport.Cells(outputRow, j + 4 + colOffset).Value = 0
                End If
            Next j

            ' Сумма по метрикам (столбцы 4–28)
            Dim itogRange As String
            itogRange = wsReport.Cells(outputRow, 4 + colOffset).Address(False, False) & ":" & _
                        wsReport.Cells(outputRow, 28 + colOffset).Address(False, False)
            wsReport.Cells(outputRow, 54 + colOffset).Formula = "=SUM(" & itogRange & ")"

            outputRow = outputRow + 1
NextEmployee:
        Next i

        lastDataRow = outputRow - 1

        If lastDataRow >= 4 Then
            With wsReport.Range(wsReport.Cells(4, 1 + colOffset), wsReport.Cells(lastDataRow, 54 + colOffset))
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With

            wsReport.Columns(3 + colOffset).NumberFormat = "0"
            wsReport.Range(wsReport.Cells(4, 4 + colOffset), wsReport.Cells(lastDataRow, 28 + colOffset)).NumberFormat = "0.00%"
            wsReport.Range(wsReport.Cells(4, 29 + colOffset), wsReport.Cells(lastDataRow, 53 + colOffset)).NumberFormat = "#,##0"
            wsReport.Range(wsReport.Cells(4, 54 + colOffset), wsReport.Cells(lastDataRow, 54 + colOffset)).NumberFormat = "0.00%"

            ' Сортировка по "Итогу"
            With wsReport.Sort
                .SortFields.Clear
                .SortFields.Add Key:=wsReport.Range(wsReport.Cells(4, 54 + colOffset), wsReport.Cells(lastDataRow, 54 + colOffset)), _
                    SortOn:=xlSortOnValues, Order:=xlDescending
                .SetRange wsReport.Range(wsReport.Cells(4, 1 + colOffset), wsReport.Cells(lastDataRow, 54 + colOffset))
                .Header = xlNo
                .Apply
            End With

            ' Нумерация
            Dim numRow As Long
            For numRow = 4 To lastDataRow
                wsReport.Cells(numRow, 1 + colOffset).Value = numRow - 3
            Next numRow

            ' Общий итог
            Dim totalRow As Long
            totalRow = lastDataRow + 2

            With wsReport.Range(wsReport.Cells(totalRow, 1 + colOffset), wsReport.Cells(totalRow, 54 + colOffset))
                .Font.Bold = True
                .Font.Size = 16
                .Interior.Color = RGB(221, 235, 247)
                .Borders(xlEdgeTop).LineStyle = xlDouble
                .Cells(1, 1).Value = "Общий итог"
                .Cells(1, 1).HorizontalAlignment = xlLeft
            End With

            Dim avgRange As String
            avgRange = wsReport.Cells(4, 54 + colOffset).Address(False, False) & ":" & _
                       wsReport.Cells(lastDataRow, 54 + colOffset).Address(False, False)
            wsReport.Cells(totalRow, 54 + colOffset).FormulaLocal = "=СРЗНАЧ(" & avgRange & ")"
            wsReport.Cells(totalRow, 54 + colOffset).NumberFormat = "0.00%"
            wsReport.Columns(54 + colOffset).ColumnWidth = 15

            '  форматирование
            Dim fmtCol As Long
            For fmtCol = 4 To 28
                With wsReport.Range(wsReport.Cells(4, fmtCol + colOffset), wsReport.Cells(lastDataRow, fmtCol + colOffset))
                    .FormatConditions.Delete
                    .FormatConditions.AddColorScale ColorScaleType:=3
                    With .FormatConditions(1).ColorScaleCriteria(1)
                        .Type = xlConditionValueLowestValue
                        .FormatColor.Color = RGB(255, 100, 100)
                    End With
                    With .FormatConditions(1).ColorScaleCriteria(2)
                        .Type = xlConditionValuePercentile
                        .Value = 50
                        .FormatColor.Color = RGB(255, 255, 100)
                    End With
                    With .FormatConditions(1).ColorScaleCriteria(3)
                        .Type = xlConditionValueHighestValue
                        .FormatColor.Color = RGB(100, 255, 100)
                    End With
                End With
            Next fmtCol

            With wsReport.Range(wsReport.Cells(4, 54 + colOffset), wsReport.Cells(lastDataRow, 54 + colOffset))
                .FormatConditions.Delete
                .FormatConditions.AddColorScale ColorScaleType:=3
                With .FormatConditions(1).ColorScaleCriteria(1)
                    .Type = xlConditionValueLowestValue
                    .FormatColor.Color = RGB(255, 100, 100)
                End With
                With .FormatConditions(1).ColorScaleCriteria(2)
                    .Type = xlConditionValuePercentile
                    .Value = 50
                    .FormatColor.Color = RGB(255, 255, 100)
                End With
                With .FormatConditions(1).ColorScaleCriteria(3)
                    .Type = xlConditionValueHighestValue
                    .FormatColor.Color = RGB(100, 255, 100)
                End With
            End With

            ' скрытие
            wsReport.Range(wsReport.Columns(2 + colOffset), wsReport.Columns(53 + colOffset)).Hidden = True
            wsReport.Columns(2 + colOffset).AutoFit

            '  рамка
            With wsReport.Range(wsReport.Cells(3, 1 + colOffset), wsReport.Cells(totalRow, 54 + colOffset))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                .Borders.Weight = xlThin
            End With
        End If
    Next reportType

        ' === ЛИСТ "Выработка по потокам" ===
    Dim wsFlow As Worksheet
    On Error Resume Next
    Set wsFlow = ActiveWorkbook.Sheets("Выработка по потокам")
    If Not wsFlow Is Nothing Then wsFlow.Delete
    On Error GoTo 0

    Set wsFlow = ActiveWorkbook.Sheets.Add(After:=wsReport)
    wsFlow.Name = "Выработка по потокам"

    Dim pasteCol As Long: pasteCol = 1

    wsReport.Columns("DI:DJ").Copy
    wsFlow.Columns(pasteCol).PasteSpecial Paste:=xlPasteValues
    pasteCol = pasteCol + 2

    wsReport.Columns("DL:EJ").Copy
    wsFlow.Columns(pasteCol).PasteSpecial Paste:=xlPasteValues
    pasteCol = pasteCol + 25

    wsReport.Columns("FJ").Copy
    wsFlow.Columns(pasteCol).PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    wsFlow.Cells.EntireColumn.Hidden = False
    wsFlow.Cells.EntireRow.Hidden = False
    '  формат
    wsReport.Range("DL4").Copy
    For col = 3 To 28
        wsFlow.Columns(col).PasteSpecial Paste:=xlPasteFormats
    Next col
    Application.CutCopyMode = False

    wsReport.Range("DI3:EJ3").Copy
    wsFlow.Range("A3:AB3").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsFlow.Rows(3).WrapText = True
    wsFlow.Rows(3).AutoFilter

    Dim refWidth As Double
    refWidth = wsFlow.Columns("AB").ColumnWidth
    wsFlow.Range("C:Z").ColumnWidth = refWidth

    wsFlow.Columns("A:B").AutoFit
    wsFlow.Columns("AB:XFD").AutoFit

    ' === СВОД ПО ПОТОКАМ ===
    Dim summaryRow As Long
    summaryRow = wsReport.Cells(wsReport.Rows.Count, "CC").End(xlUp).Row + 2

    wsReport.Cells(3, 169).Value = "Отбор"
    wsReport.Cells(4, 169).Value = "Приемка"
    wsReport.Cells(5, 169).Value = "Размещение"
    wsReport.Cells(6, 169).Value = "Упаковка"

    With wsReport.Range(wsReport.Cells(3, 169), wsReport.Cells(6, 169))
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    Dim dataStartRow As Long, dataEndRow As Long
    dataStartRow = 4
    dataEndRow = wsReport.Cells(wsReport.Rows.Count, "DJ").End(xlUp).Row

    Dim formulaOtbor As String
    formulaOtbor = "=AVERAGE(DL" & dataStartRow & ":DL" & dataEndRow & ",DM" & dataStartRow & ":DM" & dataEndRow & ",DN" & dataStartRow & ":DN" & dataEndRow & ",DO" & dataStartRow & ":DO" & dataEndRow & ",DP" & dataStartRow & ":DP" & dataEndRow & ")"
    wsReport.Cells(3, 170).Formula = formulaOtbor

    Dim formulaPriemka As String
    formulaPriemka = "=AVERAGE(DS" & dataStartRow & ":DS" & dataEndRow & ",DT" & dataStartRow & ":DT" & dataEndRow & ")"
    wsReport.Cells(3 + 1, 170).Formula = formulaPriemka

    Dim formulaRazmeshenie As String
    formulaRazmeshenie = "=AVERAGE(DU" & dataStartRow & ":DU" & dataEndRow & ",DV" & dataStartRow & ":DV" & dataEndRow & ")"
    wsReport.Cells(3 + 2, 170).Formula = formulaRazmeshenie

    Dim formulaUpakovka As String
    formulaUpakovka = "=AVERAGE(EB" & dataStartRow & ":EB" & dataEndRow & ",ED" & dataStartRow & ":ED" & dataEndRow & ",EE" & dataStartRow & ":EE" & dataEndRow & ")"
    wsReport.Cells(3 + 3, 170).Formula = formulaUpakovka
    
    Dim svodPoPotokam As String
    svodPoPotokam = "=SUM(FN3:FN6)"
    wsReport.Cells(7, 170).Formula = svodPoPotokam
    
    With wsReport.Range(wsReport.Cells(3, 170), wsReport.Cells(3 + 3, 170))
        .NumberFormat = "0.00%"
        .HorizontalAlignment = xlRight
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With

    ' Скрытие промежуточных столбцов
    wsReport.Range(wsReport.Columns(3), wsReport.Columns(53)).Hidden = True
    wsReport.Range(wsReport.Columns(59), wsReport.Columns(109)).Hidden = True
    wsReport.Range(wsReport.Columns(115), wsReport.Columns(165)).Hidden = True

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    wsReport.Activate
    wsReport.Range("A1:FN90").Copy
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.Description & vbCrLf & "В процедуре GenerateAllReportsSideBySide", vbCritical
    Resume CleanUp
End Sub

Function SheetExists(sheetName As String, Optional wb As Workbook = Nothing) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    SheetExists = (wb.Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function

Function CreateOrClearReportSheet(sheetName As String, Optional wb As Workbook = Nothing) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set CreateOrClearReportSheet = ws
End Function



