Attribute VB_Name = "ModuleDates"
' ======================================================================
' Модуль: ModuleDates
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' Назначение: Расширенная утилита для парсинга дат различных форматов
' Версия: 1.4.2 — 11.09.2025
' ======================================================================
Option Explicit

'=====================================================================
' УТИЛИТА: TryParseDateExtended
' Назначение: превратить строку почти любого «русского» формата даты
'             в реальную Date. Возвращает True, если удалось.
' Примечание: Переименована во избежание конфликта с ModuleMain
'=====================================================================
Public Function TryParseDateExtended(rawText As String, ByRef outDate As Date) As Boolean
    Dim t As String, parsedDateStr As String, ok As Boolean
    TryParseDateExtended = False
    
    If Len(Trim(rawText)) = 0 Then Exit Function
    
    On Error Resume Next
    
    ' --- 1. Прямая проверка VBA ------------------------------
    If IsDate(rawText) Then
        outDate = CDate(rawText)
        TryParseDateExtended = True
        Exit Function
    End If
    
    ' --- 2. Заменяем / на . ----------------------------------
    t = Replace(rawText, "/", ".")
    If IsDate(t) Then
        outDate = CDate(t)
        TryParseDateExtended = True
        Exit Function
    End If
    
    ' --- 3. Попытки DDMMYYYY / DDMMYY / DDMYY / DDMM ---------
    Dim clean As String, i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If IsNumeric(ch) Then clean = clean & ch
    Next i
    
    Select Case Len(clean)
        Case 8  ' 25072025 -> 25.07.2025
            parsedDateStr = Left(clean, 2) & "." & Mid(clean, 3, 2) & "." & Right(clean, 4)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
            
        Case 6  ' 250725 -> 25.07.2025
            parsedDateStr = Left(clean, 2) & "." & Mid(clean, 3, 2) & ".20" & Right(clean, 2)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
            
        Case 5  ' 25725 -> 2.57.25 (некорректно) или 25.7.25 -> попробуем разные варианты
            ' Вариант 1: D.MM.YY
            parsedDateStr = Left(clean, 1) & "." & Mid(clean, 2, 2) & ".20" & Right(clean, 2)
            ok = IsDate(parsedDateStr)
            If ok Then
                outDate = CDate(parsedDateStr)
            Else
                ' Вариант 2: DD.M.YY
                parsedDateStr = Left(clean, 2) & "." & Mid(clean, 3, 1) & ".20" & Right(clean, 2)
                ok = IsDate(parsedDateStr)
                If ok Then outDate = CDate(parsedDateStr)
            End If
            
        Case 4  ' 2507 -> 25.07.текущий_год
            parsedDateStr = Left(clean, 2) & "." & Right(clean, 2) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
            
        Case 3  ' 257 -> 2.5.текущий_год (DD.M) или 25.текущий_месяц.текущий_год (DD)
            ' Вариант 1: D.M.текущий_год
            parsedDateStr = Left(clean, 1) & "." & Mid(clean, 2, 1) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then
                outDate = CDate(parsedDateStr)
            Else
                ' Вариант 2: DD.текущий_месяц.текущий_год
                parsedDateStr = Left(clean, 2) & "." & Month(Date) & "." & Year(Date)
                ok = IsDate(parsedDateStr)
                If ok Then outDate = CDate(parsedDateStr)
            End If
            
        Case 2  ' 25 -> 25.текущий_месяц.текущий_год
            parsedDateStr = clean & "." & Month(Date) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
            
        Case 1  ' 5 -> 5.текущий_месяц.текущий_год
            parsedDateStr = clean & "." & Month(Date) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
            
        Case Else
            ok = False
    End Select
    
    TryParseDateExtended = ok
    On Error GoTo 0
End Function

'=====================================================================
' УТИЛИТА: FormatDateRussian
' Назначение: Форматирование даты в русском стиле (ИСПРАВЛЕНО)
'=====================================================================
Public Function FormatDateRussian(inputDate As Date, Optional style As String = "full") As String
    On Error GoTo FormatError
    
    Select Case LCase(style)
        Case "full"
            ' ИСПРАВЛЕНИЕ: Используем прямое обращение к месяцам
            Dim dayNum As Integer, monthNum As Integer, yearNum As Integer
            dayNum = Day(inputDate)
            monthNum = Month(inputDate)
            yearNum = Year(inputDate)
            
            FormatDateRussian = "«" & dayNum & "» " & GetDirectRussianMonth(monthNum) & " " & yearNum & " г."
            
        Case "short"
            FormatDateRussian = Format(inputDate, "dd.mm.yyyy")
            
        Case "medium"
            FormatDateRussian = Day(inputDate) & " " & GetMonthNameShortRussian(Month(inputDate)) & " " & Year(inputDate)
            
        Case Else
            FormatDateRussian = Format(inputDate, "dd.mm.yyyy")
    End Select
    
    Exit Function
    
FormatError:
    FormatDateRussian = Format(inputDate, "dd.mm.yyyy")
End Function

' НОВАЯ ФУНКЦИЯ в ModuleDates - добавьте её
Private Function GetDirectRussianMonth(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetDirectRussianMonth = "января"
        Case 2: GetDirectRussianMonth = "февраля"
        Case 3: GetDirectRussianMonth = "марта"
        Case 4: GetDirectRussianMonth = "апреля"
        Case 5: GetDirectRussianMonth = "мая"
        Case 6: GetDirectRussianMonth = "июня"
        Case 7: GetDirectRussianMonth = "июля"
        Case 8: GetDirectRussianMonth = "августа"
        Case 9: GetDirectRussianMonth = "сентября"    ' < ЗДЕСЬ ГЛАВНОЕ
        Case 10: GetDirectRussianMonth = "октября"
        Case 11: GetDirectRussianMonth = "ноября"
        Case 12: GetDirectRussianMonth = "декабря"
        Case Else: GetDirectRussianMonth = "неизвестно"
    End Select
End Function


' НОВАЯ ФУНКЦИЯ: Прямое получение названий месяцев
Private Function GetRussianMonthDirectly(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetRussianMonthDirectly = "января"
        Case 2: GetRussianMonthDirectly = "февраля"
        Case 3: GetRussianMonthDirectly = "марта"
        Case 4: GetRussianMonthDirectly = "апреля"
        Case 5: GetRussianMonthDirectly = "мая"
        Case 6: GetRussianMonthDirectly = "июня"
        Case 7: GetRussianMonthDirectly = "июля"
        Case 8: GetRussianMonthDirectly = "августа"
        Case 9: GetRussianMonthDirectly = "сентября"  ' < КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ
        Case 10: GetRussianMonthDirectly = "октября"
        Case 11: GetRussianMonthDirectly = "ноября"
        Case 12: GetRussianMonthDirectly = "декабря"
        Case Else: GetRussianMonthDirectly = "неизвестно"
    End Select
End Function



'=====================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
'=====================================================================
'Private Function GetMonthNameRussian(monthNumber As Integer) As String
'    Select Case monthNumber
'        Case 1: GetMonthNameRussian = "января"
'        Case 2: GetMonthNameRussian = "февраля"
'        Case 3: GetMonthNameRussian = "марта"
'        Case 4: GetMonthNameRussian = "апреля"
'        Case 5: GetMonthNameRussian = "мая"
'        Case 6: GetMonthNameRussian = "июня"
'        Case 7: GetMonthNameRussian = "июля"
'        Case 8: GetMonthNameRussian = "августа"
'        Case 9: GetMonthNameRussian = "сентября"
'        Case 10: GetMonthNameRussian = "октября"
'        Case 11: GetMonthNameRussian = "ноября"
'        Case 12: GetMonthNameRussian = "декабря"
'        Case Else: GetMonthNameRussian = "неизвестно"
'    End Select
'End Function

' НОВАЯ ФУНКЦИЯ: Родительный падеж для правильных окончаний
Private Function GetMonthNameRussianGenitive(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetMonthNameRussianGenitive = "января"
        Case 2: GetMonthNameRussianGenitive = "февраля"
        Case 3: GetMonthNameRussianGenitive = "марта"
        Case 4: GetMonthNameRussianGenitive = "апреля"
        Case 5: GetMonthNameRussianGenitive = "мая"
        Case 6: GetMonthNameRussianGenitive = "июня"
        Case 7: GetMonthNameRussianGenitive = "июля"
        Case 8: GetMonthNameRussianGenitive = "августа"
        Case 9: GetMonthNameRussianGenitive = "сентября"
        Case 10: GetMonthNameRussianGenitive = "октября"
        Case 11: GetMonthNameRussianGenitive = "ноября"
        Case 12: GetMonthNameRussianGenitive = "декабря"
        Case Else: GetMonthNameRussianGenitive = "неизвестно"
    End Select
End Function

Private Function GetMonthNameShortRussian(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetMonthNameShortRussian = "янв"
        Case 2: GetMonthNameShortRussian = "фев"
        Case 3: GetMonthNameShortRussian = "мар"
        Case 4: GetMonthNameShortRussian = "апр"
        Case 5: GetMonthNameShortRussian = "май"
        Case 6: GetMonthNameShortRussian = "июн"
        Case 7: GetMonthNameShortRussian = "июл"
        Case 8: GetMonthNameShortRussian = "авг"
        Case 9: GetMonthNameShortRussian = "сен"
        Case 10: GetMonthNameShortRussian = "окт"
        Case 11: GetMonthNameShortRussian = "ноя"
        Case 12: GetMonthNameShortRussian = "дек"
        Case Else: GetMonthNameShortRussian = "???"
    End Select
End Function

'=====================================================================
' УТИЛИТА: ValidateDateRange
' Назначение: Проверка попадания даты в заданный диапазон
'=====================================================================
Public Function ValidateDateRange(inputDate As Date, minDate As Date, maxDate As Date) As Boolean
    ValidateDateRange = (inputDate >= minDate And inputDate <= maxDate)
End Function

'=====================================================================
' УТИЛИТА: GetWorkingDaysCount
' Назначение: Подсчет рабочих дней между датами (исключая выходные)
'=====================================================================
Public Function GetWorkingDaysCount(startDate As Date, endDate As Date) As Long
    Dim currentDate As Date, count As Long
    currentDate = startDate
    count = 0
    
    While currentDate <= endDate
        ' Проверяем, что день не суббота (7) и не воскресенье (1)
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            count = count + 1
        End If
        currentDate = currentDate + 1
    Wend
    
    GetWorkingDaysCount = count
End Function

