Attribute VB_Name = "ModuleRepository"
' ======================================================================
' Module: ModuleRepository
' Author: CreateLetter contributors
' Purpose: Workbook CRUD/search/export helpers with typed history DTO support
' Version: 1.0.2 - 29.03.2026
' ======================================================================

Option Explicit

Public Function RepositorySearchAddresses(searchTerm As String) As Collection
    Set RepositorySearchAddresses = New Collection

    On Error GoTo SearchError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")

    Dim addressData As Variant
    addressData = RepositoryReadWorksheetMatrix(ws, AddressColumnAddressee, AddressColumnPhone, AddressesTableName)
    If IsEmpty(addressData) Then Exit Function

    Dim startRow As Long
    startRow = RepositoryGetStructuredDataStartRow(ws, AddressColumnAddressee, AddressColumnPhone, AddressesTableName)

    Dim normalizedSearch As String
    normalizedSearch = UCase$(Trim$(searchTerm))

    Dim i As Long
    For i = LBound(addressData, 1) To UBound(addressData, 1)
        If Len(normalizedSearch) = 0 Or InStr(1, UCase$(BuildAddressSearchLineFromMatrix(addressData, i)), normalizedSearch, vbTextCompare) > 0 Then
            RepositorySearchAddresses.Add BuildAddressListItemFromMatrix(addressData, i, startRow + i - 1)
        End If
    Next i

    Exit Function

SearchError:
    Debug.Print "RepositorySearchAddresses error: " & Err.Description
End Function

Public Function RepositorySearchAttachments(searchTerm As String) As Collection
    Set RepositorySearchAttachments = New Collection

    On Error GoTo SearchError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")

    Dim settingsData As Variant
    settingsData = RepositoryReadWorksheetMatrix(ws, SettingsColumnAttachmentName, SettingsColumnExecutorPhone)
    If IsEmpty(settingsData) Then Exit Function

    Dim normalizedSearch As String
    normalizedSearch = UCase$(Trim$(searchTerm))

    Dim i As Long
    For i = LBound(settingsData, 1) To UBound(settingsData, 1)
        If Len(Trim$(RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnAttachmentName))) > 0 Then
            If Len(normalizedSearch) = 0 Or InStr(1, UCase$(RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnAttachmentName)), normalizedSearch, vbTextCompare) > 0 Then
                RepositorySearchAttachments.Add RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnAttachmentName)
            End If
        End If
    Next i

    Exit Function

SearchError:
    Debug.Print "RepositorySearchAttachments error: " & Err.Description
End Function

Public Function RepositoryGetExecutorsList() As Collection
    Set RepositoryGetExecutorsList = New Collection

    On Error GoTo LookupError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")

    Dim settingsData As Variant
    settingsData = RepositoryReadWorksheetMatrix(ws, SettingsColumnAttachmentName, SettingsColumnExecutorPhone)
    If IsEmpty(settingsData) Then Exit Function

    Dim i As Long
    For i = LBound(settingsData, 1) To UBound(settingsData, 1)
        If Len(Trim$(RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnExecutorName))) > 0 Then
            RepositoryGetExecutorsList.Add RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnExecutorName)
        End If
    Next i

    Exit Function

LookupError:
    Debug.Print "RepositoryGetExecutorsList error: " & Err.Description
End Function

Public Function RepositoryGetExecutorPhone(executorFIO As String) As String
    RepositoryGetExecutorPhone = t("common.not_specified", "Not specified")

    On Error GoTo LookupError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")

    Dim settingsData As Variant
    settingsData = RepositoryReadWorksheetMatrix(ws, SettingsColumnAttachmentName, SettingsColumnExecutorPhone)
    If IsEmpty(settingsData) Then Exit Function

    Dim i As Long
    For i = LBound(settingsData, 1) To UBound(settingsData, 1)
        If RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnExecutorName) = executorFIO Then
            If Len(Trim$(RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnExecutorPhone))) > 0 Then
                RepositoryGetExecutorPhone = RepositoryMatrixValueOrEmpty(settingsData, i, SettingsColumnExecutorPhone)
            End If
            Exit Function
        End If
    Next i

    Exit Function

LookupError:
    Debug.Print "RepositoryGetExecutorPhone error: " & Err.Description
End Function

Public Sub RepositorySaveNewAddress(addressArray As Variant)
    On Error GoTo SaveError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")

    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.Count, AddressColumnAddressee).End(xlUp).Row + 1
    WriteAddressRow ws, newRow, addressArray
    Exit Sub

SaveError:
    MsgBox t("core.address.error.save", "Error saving address: ") & Err.Description, vbCritical
End Sub

Public Sub RepositoryUpdateExistingAddress(rowNumber As Long, addressArray As Variant)
    On Error GoTo UpdateError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")

    WriteAddressRow ws, rowNumber, addressArray
    Exit Sub

UpdateError:
    MsgBox t("core.address.error.update", "Error updating address: ") & Err.Description, vbCritical
End Sub

Public Sub RepositoryDeleteExistingAddress(rowNumber As Long)
    On Error GoTo DeleteError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")

    ws.Rows(rowNumber).Delete
    Exit Sub

DeleteError:
    MsgBox t("core.address.error.delete", "Error deleting address: ") & Err.Description, vbCritical
End Sub

Public Function RepositoryIsAddressDuplicate(addressArray As Variant, Optional excludeRow As Long = 0) As Boolean
    RepositoryIsAddressDuplicate = False

    On Error GoTo CheckError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")

    Dim addressData As Variant
    addressData = RepositoryReadWorksheetMatrix(ws, AddressColumnAddressee, AddressColumnPhone, AddressesTableName)
    If IsEmpty(addressData) Then Exit Function

    Dim startRow As Long
    startRow = RepositoryGetStructuredDataStartRow(ws, AddressColumnAddressee, AddressColumnPhone, AddressesTableName)

    Dim i As Long
    Dim matchCount As Integer
    For i = LBound(addressData, 1) To UBound(addressData, 1)
        If startRow + i - 1 = excludeRow Then GoTo NextRow

        matchCount = 0

        If UCase$(Trim$(RepositoryMatrixValueOrEmpty(addressData, i, AddressColumnAddressee))) = UCase$(Trim$(CStr(addressArray(AddressIndexAddressee)))) Then matchCount = matchCount + 1
        If UCase$(Trim$(RepositoryMatrixValueOrEmpty(addressData, i, AddressColumnCity))) = UCase$(Trim$(CStr(addressArray(AddressIndexCity)))) Then matchCount = matchCount + 1
        If UCase$(Trim$(RepositoryMatrixValueOrEmpty(addressData, i, AddressColumnPostalCode))) = UCase$(Trim$(CStr(addressArray(AddressIndexPostalCode)))) Then matchCount = matchCount + 1

        If matchCount >= 3 Then
            RepositoryIsAddressDuplicate = True
            Exit Function
        End If

NextRow:
    Next i

    Exit Function

CheckError:
    RepositoryIsAddressDuplicate = False
End Function

Public Function RepositoryLoadLetterHistoryData() As Collection
    Set RepositoryLoadLetterHistoryData = New Collection

    On Error GoTo LoadError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Letters")

    Dim letterData As Variant
    letterData = RepositoryReadWorksheetMatrix(ws, LetterColumnAddressee, LetterColumnDocumentType, LettersTableName)
    If IsEmpty(letterData) Then Exit Function

    Dim startRow As Long
    startRow = RepositoryGetStructuredDataStartRow(ws, LetterColumnAddressee, LetterColumnDocumentType, LettersTableName)

    Dim i As Long
    For i = LBound(letterData, 1) To UBound(letterData, 1)
        RepositoryLoadLetterHistoryData.Add CreateLetterHistoryRecordFromMatrix(letterData, i, startRow + i - 1)
    Next i

    Exit Function

LoadError:
    Set RepositoryLoadLetterHistoryData = New Collection
End Function

Public Function RepositoryFilterLetterHistoryRecords(allLettersData As Collection, searchText As String) As Collection
    Set RepositoryFilterLetterHistoryRecords = New Collection

    If allLettersData Is Nothing Then Exit Function

    Dim normalizedSearch As String
    normalizedSearch = Trim$(searchText)

    Dim i As Long
    For i = 1 To allLettersData.Count
        Dim record As clsLetterHistoryRecord
        If TryGetLetterHistoryRecord(allLettersData(i), record) Then
            If Len(normalizedSearch) = 0 Or LetterHistoryRecordMatches(record, normalizedSearch) Then
                RepositoryFilterLetterHistoryRecords.Add record
            End If
        End If
    Next i
End Function

Public Function RepositoryFormatLetterHistoryDisplay(letterData As Variant) As String
    Dim record As clsLetterHistoryRecord
    If Not TryGetLetterHistoryRecord(letterData, record) Then
        RepositoryFormatLetterHistoryDisplay = CStr(letterData)
        Exit Function
    End If

    Dim formattedDate As String
    formattedDate = FormatHistoryDateForDisplay(record.OutgoingDate)

    Dim formattedSum As String
    formattedSum = FormatHistoryDocumentSum(record.DocumentSum)

    Dim statusText As String
    statusText = BuildHistoryStatusLabel(record.ReturnStatus)

    Dim addressee As String
    Dim attachments As String
    addressee = Left$(record.Addressee, 25) & IIf(Len(record.Addressee) > 25, "...", "")
    attachments = Left$(record.AttachmentText, 30) & IIf(Len(record.AttachmentText) > 30, "...", "")

    RepositoryFormatLetterHistoryDisplay = addressee & " | " & _
                                           record.OutgoingNumber & " | " & _
                                           formattedDate & " | " & _
                                           attachments & " | " & _
                                           formattedSum & " | " & _
                                           statusText & " | " & _
                                           record.Executor & " | " & _
                                           GetDocumentTypeDisplayLabel(record.DocumentTypeKey)
End Function

Public Function RepositoryTryParseLetterHistoryRecord(letterData As Variant, ByRef parts As Variant) As Boolean
    Dim record As clsLetterHistoryRecord
    If TryGetLetterHistoryRecord(letterData, record) Then
        parts = record.ToPartsArray()
        RepositoryTryParseLetterHistoryRecord = True
        Exit Function
    End If

    parts = Empty
    RepositoryTryParseLetterHistoryRecord = False
End Function

Public Function RepositoryBuildLetterReturnStatus(isReceived As Boolean, returnDateText As String) As String
    If isReceived Then
        RepositoryBuildLetterReturnStatus = Format$(ResolveLetterDateOrToday(returnDateText), "dd.mm.yyyy") & t("history.status.received_suffix", " получено")
    Else
        RepositoryBuildLetterReturnStatus = t("history.status.not_received", "не получено")
    End If
End Function

Public Function RepositoryGetLetterHistorySearchHintsText() As String
    RepositoryGetLetterHistorySearchHintsText = t("form.letter_history.msg.search_hints_body", _
                                                  "ПОДСКАЗКИ ПО ПОИСКУ:" & vbCrLf & vbCrLf & _
                                                  "• Для поиска по сумме вводите только цифры: 125000" & vbCrLf & _
                                                  "• Система найдет '125000', '125 000', '125000 руб.'" & vbCrLf & _
                                                  "• Поиск работает сразу по всем колонкам" & vbCrLf & _
                                                  "• Можно искать по части слова или номера" & vbCrLf & vbCrLf & _
                                                  "Если вы вручную меняли Excel, нажмите 'Обновить данные'")
End Function

Public Sub RepositoryExportLetterHistoryRecords(records As Collection)
    If records Is Nothing Or records.Count = 0 Then
        MsgBox t("form.letter_history.msg.no_export_data", "No data to export."), vbExclamation
        Exit Sub
    End If

    On Error GoTo ExportError

    Dim exportWb As Workbook
    Dim exportWs As Worksheet
    Set exportWb = Workbooks.Add
    Set exportWs = exportWb.Worksheets(1)

    WriteLetterHistoryExportHeaders exportWs
    WriteLetterHistoryExportRecords exportWs, records

    exportWs.Columns("A:H").AutoFit
    exportWs.Name = t("form.letter_history.export.sheet_name", "Letters history ") & Format$(Date, "dd.mm.yyyy")
    exportWb.Application.Visible = True

    MsgBox t("form.letter_history.msg.export_completed", "Export completed.") & vbCrLf & _
           t("form.letter_history.msg.records_exported", "Records exported: ") & records.Count, _
           vbInformation, _
           t("form.letter_history.msg.export_title", "Data export")
    Exit Sub

ExportError:
    MsgBox t("form.letter_history.msg.export_error", "Export error: ") & Err.Description, vbCritical
End Sub

Public Function RepositoryHasReturnStatusDate(returnStatus As String) As Boolean
    RepositoryHasReturnStatusDate = (Len(RepositoryExtractReturnStatusDate(returnStatus)) > 0)
End Function

Public Function RepositoryExtractReturnStatusDate(returnStatus As String) As String
    Dim parts() As String
    Dim index As Long

    RepositoryExtractReturnStatusDate = ""
    parts = Split(returnStatus, " ")

    For index = LBound(parts) To UBound(parts)
        If IsDate(parts(index)) Then
            RepositoryExtractReturnStatusDate = Format$(CDate(parts(index)), "dd.mm.yyyy")
            Exit Function
        End If
    Next index
End Function

Public Sub RepositoryUpdateLetterHistoryRow(rowNumber As Long, sumValue As String, returnStatus As String)
    On Error GoTo UpdateError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Letters")

    If Len(Trim$(sumValue)) = 0 Then
        ws.Cells(rowNumber, LetterColumnDocumentSum).Value = ""
    ElseIf IsNumeric(sumValue) Then
        ws.Cells(rowNumber, LetterColumnDocumentSum).Value = CDbl(sumValue)
    Else
        ws.Cells(rowNumber, LetterColumnDocumentSum).Value = sumValue
    End If

    ws.Cells(rowNumber, LetterColumnReturnStatus).Value = returnStatus
    Exit Sub

UpdateError:
    Err.Raise Err.Number, "RepositoryUpdateLetterHistoryRow", Err.Description
End Sub

Public Function RepositoryGetStructuredDataRange(ws As Worksheet, firstColumn As Long, lastColumn As Long, Optional preferredTableName As String = "") As Range
    On Error GoTo RangeError

    If Len(Trim$(preferredTableName)) > 0 Then
        Dim preferredTable As ListObject
        If TryGetWorksheetTable(ws, preferredTableName, preferredTable) Then
            If Not preferredTable.DataBodyRange Is Nothing Then
                Set RepositoryGetStructuredDataRange = preferredTable.DataBodyRange
                Exit Function
            End If
        End If
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, firstColumn).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Function

    Set RepositoryGetStructuredDataRange = ws.Range(ws.Cells(FIRST_DATA_ROW, firstColumn), ws.Cells(lastRow, lastColumn))
    Exit Function

RangeError:
    Set RepositoryGetStructuredDataRange = Nothing
End Function

Public Function RepositoryReadWorksheetMatrix(ws As Worksheet, firstColumn As Long, lastColumn As Long, Optional preferredTableName As String = "") As Variant
    Dim sourceRange As Range
    Set sourceRange = RepositoryGetStructuredDataRange(ws, firstColumn, lastColumn, preferredTableName)

    If sourceRange Is Nothing Then
        RepositoryReadWorksheetMatrix = Empty
        Exit Function
    End If

    RepositoryReadWorksheetMatrix = sourceRange.Value
End Function

Public Function RepositoryGetStructuredDataStartRow(ws As Worksheet, firstColumn As Long, lastColumn As Long, Optional preferredTableName As String = "") As Long
    Dim sourceRange As Range
    Set sourceRange = RepositoryGetStructuredDataRange(ws, firstColumn, lastColumn, preferredTableName)

    If sourceRange Is Nothing Then
        RepositoryGetStructuredDataStartRow = FIRST_DATA_ROW
    Else
        RepositoryGetStructuredDataStartRow = sourceRange.Row
    End If
End Function

Public Function RepositoryMatrixValueOrEmpty(dataMatrix As Variant, rowIndex As Long, columnIndex As Long) As String
    If IsArray(dataMatrix) Then
        RepositoryMatrixValueOrEmpty = CStr(dataMatrix(rowIndex, columnIndex))
    Else
        RepositoryMatrixValueOrEmpty = ""
    End If
End Function

Private Function TryGetWorksheetTable(ws As Worksheet, tableName As String, ByRef targetTable As ListObject) As Boolean
    On Error GoTo LookupFailed

    Set targetTable = ws.ListObjects(tableName)
    TryGetWorksheetTable = Not targetTable Is Nothing
    Exit Function

LookupFailed:
    Set targetTable = Nothing
    TryGetWorksheetTable = False
End Function

Private Function BuildAddressSearchLineFromMatrix(addressData As Variant, rowIndex As Long) As String
    BuildAddressSearchLineFromMatrix = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnAddressee) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnStreet) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnCity) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnDistrict) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnRegion) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPostalCode) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPhone)
End Function

Private Function BuildAddressListItemFromMatrix(addressData As Variant, rowIndex As Long, worksheetRowNumber As Long) As String
    BuildAddressListItemFromMatrix = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnAddressee) & " | " & _
                                     RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnStreet) & " | " & _
                                     RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnCity) & " | " & _
                                     RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnDistrict) & " | " & _
                                     RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnRegion) & " | " & _
                                     RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPostalCode) & " | " & _
                                     RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPhone) & " | " & worksheetRowNumber
End Function

Private Function CreateLetterHistoryRecordFromMatrix(letterData As Variant, rowIndex As Long, worksheetRowNumber As Long) As clsLetterHistoryRecord
    Dim record As clsLetterHistoryRecord
    Set record = New clsLetterHistoryRecord

    record.Addressee = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnAddressee)
    record.OutgoingNumber = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnOutgoingNumber)
    record.OutgoingDate = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnOutgoingDate)
    record.AttachmentText = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnAttachmentText)
    record.DocumentSum = NormalizeHistorySumCell(letterData(rowIndex, LetterColumnDocumentSum))
    record.ReturnStatus = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnReturnStatus)
    record.Executor = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnExecutor)
    record.DocumentTypeKey = NormalizeDocumentTypeKey(RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnDocumentType))
    record.RowNumber = worksheetRowNumber

    Set CreateLetterHistoryRecordFromMatrix = record
End Function

Private Function NormalizeHistorySumCell(cellValue As Variant) As String
    If IsNumeric(cellValue) And cellValue <> 0 Then
        If cellValue = Int(cellValue) Then
            NormalizeHistorySumCell = CStr(CLng(cellValue))
        Else
            NormalizeHistorySumCell = CStr(cellValue)
        End If
    Else
        NormalizeHistorySumCell = CStr(cellValue)
    End If
End Function

Private Sub WriteAddressRow(ws As Worksheet, rowNumber As Long, addressArray As Variant)
    ws.Cells(rowNumber, AddressColumnAddressee).Value = addressArray(AddressIndexAddressee)
    ws.Cells(rowNumber, AddressColumnStreet).Value = addressArray(AddressIndexStreet)
    ws.Cells(rowNumber, AddressColumnCity).Value = addressArray(AddressIndexCity)
    ws.Cells(rowNumber, AddressColumnDistrict).Value = addressArray(AddressIndexDistrict)
    ws.Cells(rowNumber, AddressColumnRegion).Value = addressArray(AddressIndexRegion)
    ws.Cells(rowNumber, AddressColumnPostalCode).Value = addressArray(AddressIndexPostalCode)
    ws.Cells(rowNumber, AddressColumnPhone).Value = FormatPhoneNumber(CStr(addressArray(AddressIndexPhone)))
End Sub

Private Function TryGetLetterHistoryRecord(letterData As Variant, ByRef record As clsLetterHistoryRecord) As Boolean
    On Error GoTo ParseFailed

    If IsObject(letterData) Then
        If TypeName(letterData) = "clsLetterHistoryRecord" Then
            Set record = letterData
            TryGetLetterHistoryRecord = True
            Exit Function
        End If
    End If

    If VarType(letterData) = vbString Then
        Set record = TryParseLegacyLetterHistoryRecord(CStr(letterData))
        TryGetLetterHistoryRecord = Not record Is Nothing
        Exit Function
    End If

ParseFailed:
    Set record = Nothing
    TryGetLetterHistoryRecord = False
End Function

Private Function TryParseLegacyLetterHistoryRecord(letterData As String) As clsLetterHistoryRecord
    On Error GoTo ParseFailed

    Dim parts() As String
    parts = Split(letterData, "|")
    If UBound(parts) < HistoryPartRowNumber Then Exit Function

    Dim record As clsLetterHistoryRecord
    Set record = New clsLetterHistoryRecord

    record.Addressee = parts(HistoryPartAddressee)
    record.OutgoingNumber = parts(HistoryPartOutgoingNumber)
    record.OutgoingDate = parts(HistoryPartOutgoingDate)
    record.AttachmentText = parts(HistoryPartAttachmentText)
    record.DocumentSum = parts(HistoryPartDocumentSum)
    record.ReturnStatus = parts(HistoryPartReturnStatus)
    record.Executor = parts(HistoryPartExecutor)
    record.DocumentTypeKey = NormalizeDocumentTypeKey(parts(HistoryPartDocumentType))
    record.RowNumber = CLng(parts(HistoryPartRowNumber))

    Set TryParseLegacyLetterHistoryRecord = record
    Exit Function

ParseFailed:
    Set TryParseLegacyLetterHistoryRecord = Nothing
End Function

Private Function LetterHistoryRecordMatches(record As clsLetterHistoryRecord, searchText As String) As Boolean
    Dim searchPattern As String
    searchPattern = UCase$(Trim$(searchText))

    If InStr(1, UCase$(record.Addressee), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(record.OutgoingNumber), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(record.OutgoingDate), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(record.AttachmentText), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If IsNumeric(searchPattern) Then
        If IsHistoryNumericMatch(record.DocumentSum, searchPattern) Then
            LetterHistoryRecordMatches = True
            Exit Function
        End If
    ElseIf InStr(1, UCase$(record.DocumentSum), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(record.ReturnStatus), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(record.Executor), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(record.DocumentTypeKey), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
        Exit Function
    End If

    If InStr(1, UCase$(GetDocumentTypeDisplayLabel(record.DocumentTypeKey)), searchPattern, vbTextCompare) > 0 Then
        LetterHistoryRecordMatches = True
    End If
End Function

Private Function IsHistoryNumericMatch(cellValue As String, searchValue As String) As Boolean
    Dim cleanCellValue As String
    Dim cleanSearchValue As String

    cleanCellValue = ExtractDigitsOnly(cellValue)
    cleanSearchValue = ExtractDigitsOnly(searchValue)

    If Len(cleanCellValue) = 0 Or Len(cleanSearchValue) = 0 Then Exit Function
    If cleanCellValue = cleanSearchValue Then
        IsHistoryNumericMatch = True
        Exit Function
    End If

    If InStr(1, cleanCellValue, cleanSearchValue, vbTextCompare) > 0 Then
        IsHistoryNumericMatch = True
    End If
End Function

Private Function ExtractDigitsOnly(inputText As String) As String
    Dim i As Long
    Dim currentChar As String

    For i = 1 To Len(inputText)
        currentChar = Mid$(inputText, i, 1)
        If currentChar >= "0" And currentChar <= "9" Then
            ExtractDigitsOnly = ExtractDigitsOnly & currentChar
        End If
    Next i
End Function

Private Sub WriteLetterHistoryExportHeaders(exportWs As Worksheet)
    With exportWs
        .Cells(1, LetterColumnAddressee).Value = t("form.letter_history.export.header.addressee", "Addressee")
        .Cells(1, LetterColumnOutgoingNumber).Value = t("form.letter_history.export.header.outgoing_number", "Outgoing Number")
        .Cells(1, LetterColumnOutgoingDate).Value = t("form.letter_history.export.header.outgoing_date", "Outgoing Date")
        .Cells(1, LetterColumnAttachmentText).Value = t("form.letter_history.export.header.attachment_name", "Attachment Name")
        .Cells(1, LetterColumnDocumentSum).Value = t("form.letter_history.export.header.document_sum", "Document Sum")
        .Cells(1, LetterColumnReturnStatus).Value = t("form.letter_history.export.header.return_mark", "Return Mark")
        .Cells(1, LetterColumnExecutor).Value = t("form.letter_history.export.header.executor_name", "Executor Name")
        .Cells(1, LetterColumnDocumentType).Value = t("form.letter_history.export.header.send_type", "Send Type")

        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.ColorIndex = 15
            .Font.ColorIndex = 1
        End With
    End With
End Sub

Private Sub WriteLetterHistoryExportRecords(exportWs As Worksheet, records As Collection)
    Dim i As Long
    For i = 1 To records.Count
        WriteLetterHistoryExportRow exportWs, i + 1, records(i)
    Next i
End Sub

Private Sub WriteLetterHistoryExportRow(exportWs As Worksheet, targetRow As Long, letterData As Variant)
    Dim record As clsLetterHistoryRecord
    If Not TryGetLetterHistoryRecord(letterData, record) Then Exit Sub

    exportWs.Cells(targetRow, LetterColumnAddressee).Value = record.Addressee
    exportWs.Cells(targetRow, LetterColumnOutgoingNumber).Value = record.OutgoingNumber
    exportWs.Cells(targetRow, LetterColumnOutgoingDate).Value = record.OutgoingDate
    exportWs.Cells(targetRow, LetterColumnAttachmentText).Value = record.AttachmentText
    exportWs.Cells(targetRow, LetterColumnDocumentSum).Value = record.DocumentSum
    exportWs.Cells(targetRow, LetterColumnReturnStatus).Value = record.ReturnStatus
    exportWs.Cells(targetRow, LetterColumnExecutor).Value = record.Executor
    exportWs.Cells(targetRow, LetterColumnDocumentType).Value = GetDocumentTypeDisplayLabel(record.DocumentTypeKey)
End Sub

Private Function FormatHistoryDateForDisplay(dateValue As String) As String
    On Error GoTo FormatFailed

    If IsDate(dateValue) Then
        FormatHistoryDateForDisplay = Format$(CDate(dateValue), "dd.mm.yyyy")
    Else
        FormatHistoryDateForDisplay = dateValue
    End If
    Exit Function

FormatFailed:
    FormatHistoryDateForDisplay = dateValue
End Function

Private Function FormatHistoryDocumentSum(sumText As String) As String
    If Len(Trim$(sumText)) > 0 And IsNumeric(sumText) Then
        If CDbl(sumText) > 0 Then
            FormatHistoryDocumentSum = Format$(CDbl(sumText), "#,##0.00") & " rub."
        Else
            FormatHistoryDocumentSum = "-"
        End If
    Else
        FormatHistoryDocumentSum = "-"
    End If
End Function

Private Function BuildHistoryStatusLabel(returnStatus As String) As String
    If RepositoryHasReturnStatusDate(returnStatus) Then
        BuildHistoryStatusLabel = t("history.status.received_label", "Received ") & returnStatus
    Else
        BuildHistoryStatusLabel = t("history.status.pending_label", "Pending ") & returnStatus
    End If
End Function
