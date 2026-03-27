Attribute VB_Name = "ModuleDates"
' ======================================================================
' Module: ModuleDates
' Author: CreateLetter contributors
' Purpose: Date parsing and Russian date-formatting helpers
' Version: 1.4.4 - 27.03.2026
' ======================================================================
Option Explicit

' Try to parse flexible date input into a VBA Date value.
Public Function TryParseDateExtended(rawText As String, ByRef outDate As Date) As Boolean
    Dim t As String
    Dim parsedDateStr As String
    Dim ok As Boolean
    Dim clean As String
    Dim i As Long
    Dim ch As String

    TryParseDateExtended = False

    If Len(Trim$(rawText)) = 0 Then Exit Function

    On Error Resume Next

    If IsDate(rawText) Then
        outDate = CDate(rawText)
        TryParseDateExtended = True
        Exit Function
    End If

    t = Replace(rawText, "/", ".")
    If IsDate(t) Then
        outDate = CDate(t)
        TryParseDateExtended = True
        Exit Function
    End If

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If IsNumeric(ch) Then clean = clean & ch
    Next i

    Select Case Len(clean)
        Case 8
            parsedDateStr = Left$(clean, 2) & "." & Mid$(clean, 3, 2) & "." & Right$(clean, 4)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
        Case 6
            parsedDateStr = Left$(clean, 2) & "." & Mid$(clean, 3, 2) & ".20" & Right$(clean, 2)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
        Case 5
            parsedDateStr = Left$(clean, 1) & "." & Mid$(clean, 2, 2) & ".20" & Right$(clean, 2)
            ok = IsDate(parsedDateStr)
            If ok Then
                outDate = CDate(parsedDateStr)
            Else
                parsedDateStr = Left$(clean, 2) & "." & Mid$(clean, 3, 1) & ".20" & Right$(clean, 2)
                ok = IsDate(parsedDateStr)
                If ok Then outDate = CDate(parsedDateStr)
            End If
        Case 4
            parsedDateStr = Left$(clean, 2) & "." & Right$(clean, 2) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
        Case 3
            parsedDateStr = Left$(clean, 1) & "." & Mid$(clean, 2, 1) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then
                outDate = CDate(parsedDateStr)
            Else
                parsedDateStr = Left$(clean, 2) & "." & Month(Date) & "." & Year(Date)
                ok = IsDate(parsedDateStr)
                If ok Then outDate = CDate(parsedDateStr)
            End If
        Case 2, 1
            parsedDateStr = clean & "." & Month(Date) & "." & Year(Date)
            ok = IsDate(parsedDateStr)
            If ok Then outDate = CDate(parsedDateStr)
        Case Else
            ok = False
    End Select

    TryParseDateExtended = ok
    On Error GoTo 0
End Function

' Format a Date value using Russian month names.
Public Function FormatDateRussian(inputDate As Date, Optional style As String = "full") As String
    On Error GoTo FormatError

    Select Case LCase$(style)
        Case "full"
            FormatDateRussian = Day(inputDate) & " " & GetDirectRussianMonth(Month(inputDate)) & " " & Year(inputDate) & " " & BuildUnicodeString(&H433, &H2E)
        Case "short"
            FormatDateRussian = Format$(inputDate, "dd.mm.yyyy")
        Case "medium"
            FormatDateRussian = Day(inputDate) & " " & GetMonthNameShortRussian(Month(inputDate)) & " " & Year(inputDate)
        Case Else
            FormatDateRussian = Format$(inputDate, "dd.mm.yyyy")
    End Select

    Exit Function

FormatError:
    FormatDateRussian = Format$(inputDate, "dd.mm.yyyy")
End Function

Private Function GetDirectRussianMonth(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetDirectRussianMonth = BuildUnicodeString(&H44F, &H43D, &H432, &H430, &H440, &H44F)
        Case 2: GetDirectRussianMonth = BuildUnicodeString(&H444, &H435, &H432, &H440, &H430, &H43B, &H44F)
        Case 3: GetDirectRussianMonth = BuildUnicodeString(&H43C, &H430, &H440, &H442, &H430)
        Case 4: GetDirectRussianMonth = BuildUnicodeString(&H430, &H43F, &H440, &H435, &H43B, &H44F)
        Case 5: GetDirectRussianMonth = BuildUnicodeString(&H43C, &H430, &H44F)
        Case 6: GetDirectRussianMonth = BuildUnicodeString(&H438, &H44E, &H43D, &H44F)
        Case 7: GetDirectRussianMonth = BuildUnicodeString(&H438, &H44E, &H43B, &H44F)
        Case 8: GetDirectRussianMonth = BuildUnicodeString(&H430, &H432, &H433, &H443, &H441, &H442, &H430)
        Case 9: GetDirectRussianMonth = BuildUnicodeString(&H441, &H435, &H43D, &H442, &H44F, &H431, &H440, &H44F)
        Case 10: GetDirectRussianMonth = BuildUnicodeString(&H43E, &H43A, &H442, &H44F, &H431, &H440, &H44F)
        Case 11: GetDirectRussianMonth = BuildUnicodeString(&H43D, &H43E, &H44F, &H431, &H440, &H44F)
        Case 12: GetDirectRussianMonth = BuildUnicodeString(&H434, &H435, &H43A, &H430, &H431, &H440, &H44F)
        Case Else: GetDirectRussianMonth = t("core.date.unknown_month", "unknown_month")
    End Select
End Function

Private Function GetRussianMonthDirectly(monthNumber As Integer) As String
    GetRussianMonthDirectly = GetDirectRussianMonth(monthNumber)
End Function

Private Function GetMonthNameRussianGenitive(monthNumber As Integer) As String
    GetMonthNameRussianGenitive = GetDirectRussianMonth(monthNumber)
End Function

Private Function GetMonthNameShortRussian(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetMonthNameShortRussian = BuildUnicodeString(&H44F, &H43D, &H432)
        Case 2: GetMonthNameShortRussian = BuildUnicodeString(&H444, &H435, &H432)
        Case 3: GetMonthNameShortRussian = BuildUnicodeString(&H43C, &H430, &H440)
        Case 4: GetMonthNameShortRussian = BuildUnicodeString(&H430, &H43F, &H440)
        Case 5: GetMonthNameShortRussian = BuildUnicodeString(&H43C, &H430, &H439)
        Case 6: GetMonthNameShortRussian = BuildUnicodeString(&H438, &H44E, &H43D)
        Case 7: GetMonthNameShortRussian = BuildUnicodeString(&H438, &H44E, &H43B)
        Case 8: GetMonthNameShortRussian = BuildUnicodeString(&H430, &H432, &H433)
        Case 9: GetMonthNameShortRussian = BuildUnicodeString(&H441, &H435, &H43D)
        Case 10: GetMonthNameShortRussian = BuildUnicodeString(&H43E, &H43A, &H442)
        Case 11: GetMonthNameShortRussian = BuildUnicodeString(&H43D, &H43E, &H44F)
        Case 12: GetMonthNameShortRussian = BuildUnicodeString(&H434, &H435, &H43A)
        Case Else: GetMonthNameShortRussian = "???"
    End Select
End Function

Public Function ValidateDateRange(inputDate As Date, minDate As Date, maxDate As Date) As Boolean
    ValidateDateRange = (inputDate >= minDate And inputDate <= maxDate)
End Function

Public Function GetWorkingDaysCount(startDate As Date, endDate As Date) As Long
    Dim currentDate As Date
    Dim workdayCount As Long

    currentDate = startDate
    workdayCount = 0

    While currentDate <= endDate
        If Weekday(currentDate) <> vbSunday And Weekday(currentDate) <> vbSaturday Then
            workdayCount = workdayCount + 1
        End If
        currentDate = currentDate + 1
    Wend

    GetWorkingDaysCount = workdayCount
End Function

Private Function BuildUnicodeString(ParamArray codePoints() As Variant) As String
    Dim i As Long

    BuildUnicodeString = ""
    For i = LBound(codePoints) To UBound(codePoints)
        BuildUnicodeString = BuildUnicodeString & ChrW$(CLng(codePoints(i)))
    Next i
End Function
