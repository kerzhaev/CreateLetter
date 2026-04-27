Attribute VB_Name = "ModuleRepository"

' ======================================================================

' Module: ModuleRepository

' Author: CreateLetter contributors

' Purpose: Workbook CRUD/search/export helpers with typed history DTO support

' Version: 1.0.5 - 29.03.2026

' ======================================================================



Option Explicit



Public Function RepositorySearchAddresses(searchTerm As String) As Collection

    Set RepositorySearchAddresses = New Collection



    On Error GoTo SearchError



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Addresses")



    Dim addressData As Variant

    addressData = RepositoryReadWorksheetMatrix(ws, AddressColumnAddressee, AddressColumnGroup, AddressesTableName)

    If IsEmpty(addressData) Then Exit Function



    Dim startRow As Long

    startRow = RepositoryGetStructuredDataStartRow(ws, AddressColumnAddressee, AddressColumnGroup, AddressesTableName)



    Dim normalizedSearch As String

    normalizedSearch = UCase$(Trim$(searchTerm))



    Dim i As Long

    For i = LBound(addressData, 1) To UBound(addressData, 1)

        If Len(normalizedSearch) = 0 Or InStr(1, UCase$(BuildAddressSearchLineFromMatrix(addressData, i)), normalizedSearch, vbTextCompare) > 0 Then

            RepositorySearchAddresses.Add CreateAddressSearchResultFromMatrix(addressData, i, startRow + i - 1)

        End If

    Next i



    Exit Function



SearchError:

    Debug.Print "RepositorySearchAddresses error: " & Err.description

End Function



Public Function RepositoryGetAddressSearchResultDisplayText(addressSearchResult As Variant) As String

    If IsAddressSearchResultArray(addressSearchResult) Then

        RepositoryGetAddressSearchResultDisplayText = CStr(addressSearchResult(AddressSearchResultDisplayText))

    Else

        RepositoryGetAddressSearchResultDisplayText = CStr(addressSearchResult)

    End If

End Function



Public Function RepositoryTryParseAddressSearchResult(addressSearchResult As Variant, ByRef addressArray As Variant, ByRef rowNumber As Long) As Boolean

    rowNumber = 0

    RepositoryTryParseAddressSearchResult = False



    If IsAddressSearchResultArray(addressSearchResult) Then

        Dim parsedAddress(AddressIndexGroup) As String



        parsedAddress(AddressIndexAddressee) = CStr(addressSearchResult(AddressSearchResultAddressee))

        parsedAddress(AddressIndexStreet) = CStr(addressSearchResult(AddressSearchResultStreet))

        parsedAddress(AddressIndexCity) = CStr(addressSearchResult(AddressSearchResultCity))

        parsedAddress(AddressIndexDistrict) = CStr(addressSearchResult(AddressSearchResultDistrict))

        parsedAddress(AddressIndexRegion) = CStr(addressSearchResult(AddressSearchResultRegion))

        parsedAddress(AddressIndexPostalCode) = CStr(addressSearchResult(AddressSearchResultPostalCode))

        parsedAddress(AddressIndexPhone) = CStr(addressSearchResult(AddressSearchResultPhone))

        parsedAddress(AddressIndexGroup) = CStr(addressSearchResult(AddressSearchResultGroup))



        addressArray = parsedAddress

        If IsNumeric(addressSearchResult(AddressSearchResultWorksheetRow)) Then

            rowNumber = CLng(addressSearchResult(AddressSearchResultWorksheetRow))

            RepositoryTryParseAddressSearchResult = (rowNumber > 0)

        End If



        Exit Function

    End If



    If VarType(addressSearchResult) = vbString Then

        Dim legacyParts As Variant

        Dim legacyErrorMessage As String

        Dim legacyRow As Long

        If TryParseAddressListItem(CStr(addressSearchResult), legacyParts, legacyRow, legacyErrorMessage) Then

            Dim legacyAddress(AddressIndexGroup) As String



            legacyAddress(AddressIndexAddressee) = CStr(legacyParts(AddressPartAddressee))

            legacyAddress(AddressIndexStreet) = CStr(legacyParts(AddressPartStreet))

            legacyAddress(AddressIndexCity) = CStr(legacyParts(AddressPartCity))

            legacyAddress(AddressIndexDistrict) = CStr(legacyParts(AddressPartDistrict))

            legacyAddress(AddressIndexRegion) = CStr(legacyParts(AddressPartRegion))

            legacyAddress(AddressIndexPostalCode) = CStr(legacyParts(AddressPartPostalCode))

            legacyAddress(AddressIndexPhone) = CStr(legacyParts(AddressPartPhone))

            legacyAddress(AddressIndexGroup) = ""



            addressArray = legacyAddress

            rowNumber = legacyRow

            RepositoryTryParseAddressSearchResult = (rowNumber > 0)

        End If

    End If

End Function



Public Function RepositoryTryLoadAddressRow(rowNumber As Long, ByRef addressArray As Variant) As Boolean

    RepositoryTryLoadAddressRow = False



    On Error GoTo LoadError



    If rowNumber < FIRST_DATA_ROW Then Exit Function



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Addresses")



    Dim loadedAddress(AddressIndexGroup) As String

    loadedAddress(AddressIndexAddressee) = CStr(ws.Cells(rowNumber, AddressColumnAddressee).value)

    loadedAddress(AddressIndexStreet) = CStr(ws.Cells(rowNumber, AddressColumnStreet).value)

    loadedAddress(AddressIndexCity) = CStr(ws.Cells(rowNumber, AddressColumnCity).value)

    loadedAddress(AddressIndexDistrict) = CStr(ws.Cells(rowNumber, AddressColumnDistrict).value)

    loadedAddress(AddressIndexRegion) = CStr(ws.Cells(rowNumber, AddressColumnRegion).value)

    loadedAddress(AddressIndexPostalCode) = CStr(ws.Cells(rowNumber, AddressColumnPostalCode).value)

    loadedAddress(AddressIndexPhone) = CStr(ws.Cells(rowNumber, AddressColumnPhone).value)

    loadedAddress(AddressIndexGroup) = CStr(ws.Cells(rowNumber, AddressColumnGroup).value)



    addressArray = loadedAddress

    RepositoryTryLoadAddressRow = True

    Exit Function



LoadError:

    RepositoryTryLoadAddressRow = False

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

    Debug.Print "RepositorySearchAttachments error: " & Err.description

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

    Debug.Print "RepositoryGetExecutorsList error: " & Err.description

End Function



Public Function RepositoryGetExecutorPhone(executorFIO As String) As String

    RepositoryGetExecutorPhone = t("common.not_specified", "Не указано")



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

    Debug.Print "RepositoryGetExecutorPhone error: " & Err.description

End Function



Public Sub RepositorySaveNewAddress(addressArray As Variant)

    On Error GoTo SaveError



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Addresses")



    Dim newRow As Long

    newRow = ws.Cells(ws.Rows.count, AddressColumnAddressee).End(xlUp).Row + 1

    WriteAddressRow ws, newRow, addressArray

    Exit Sub



SaveError:

    MsgBox t("core.address.error.save", "Ошибка при сохранении адреса: ") & Err.description, vbCritical

End Sub



Public Sub RepositoryUpdateExistingAddress(rowNumber As Long, addressArray As Variant)

    On Error GoTo UpdateError



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Addresses")



    WriteAddressRow ws, rowNumber, addressArray

    Exit Sub



UpdateError:

    MsgBox t("core.address.error.update", "Ошибка при обновлении адреса: ") & Err.description, vbCritical

End Sub



Public Sub RepositoryDeleteExistingAddress(rowNumber As Long)

    On Error GoTo DeleteError



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Addresses")



    ws.Rows(rowNumber).Delete

    Exit Sub



DeleteError:

    MsgBox t("core.address.error.delete", "Ошибка при удалении адреса: ") & Err.description, vbCritical

End Sub



Public Function RepositoryIsAddressDuplicate(addressArray As Variant, Optional excludeRow As Long = 0) As Boolean

    RepositoryIsAddressDuplicate = False



    On Error GoTo CheckError



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Addresses")



    Dim addressData As Variant

    addressData = RepositoryReadWorksheetMatrix(ws, AddressColumnAddressee, AddressColumnGroup, AddressesTableName)

    If IsEmpty(addressData) Then Exit Function



    Dim startRow As Long

    startRow = RepositoryGetStructuredDataStartRow(ws, AddressColumnAddressee, AddressColumnGroup, AddressesTableName)



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

    letterData = RepositoryReadWorksheetMatrix(ws, LetterColumnAddressee, LetterColumnDispatchRegistryDate, LettersTableName)

    If IsEmpty(letterData) Then Exit Function



    Dim startRow As Long

    startRow = RepositoryGetStructuredDataStartRow(ws, LetterColumnAddressee, LetterColumnDispatchRegistryDate, LettersTableName)



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

    For i = 1 To allLettersData.count

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

    statusText = BuildHistoryStatusLabel(record.returnStatus)



    Dim Addressee As String

    Dim attachments As String

    Addressee = Left$(record.Addressee, 25) & IIf(Len(record.Addressee) > 25, "...", "")

    attachments = Left$(record.attachmentText, 30) & IIf(Len(record.attachmentText) > 30, "...", "")



    RepositoryFormatLetterHistoryDisplay = Addressee & " | " & _
                                           record.OutgoingNumber & " | " & _
                                           formattedDate & " | " & _
                                           attachments & " | " & _
                                           formattedSum & " | " & _
                                           statusText & " | " & _
                                           record.Executor & " | " & _
                                           GetDocumentTypeDisplayLabel(record.DocumentTypeKey)

End Function

Public Function RepositoryGetLetterHistoryPackedStatusDisplay(letterData As Variant) As String

    Dim record As clsLetterHistoryRecord

    If Not TryGetLetterHistoryRecord(letterData, record) Then Exit Function

    If Len(Trim$(record.DispatchPackedFlag)) = 0 Then
        RepositoryGetLetterHistoryPackedStatusDisplay = t("history.dispatch_status.not_packed", "Нет")
        Exit Function
    End If

    If Len(Trim$(record.DispatchRegistryNumber)) > 0 Then
        RepositoryGetLetterHistoryPackedStatusDisplay = t("history.dispatch_status.registry_prefix", "Реестр") & " " & record.DispatchRegistryNumber
    Else
        RepositoryGetLetterHistoryPackedStatusDisplay = t("history.dispatch_status.packed", "Да")
    End If

End Function

Public Function RepositoryTryResolveLetterRowNumber( _
    ByVal Addressee As String, _
    ByVal OutgoingNumber As String, _
    ByVal OutgoingDate As String, _
    ByRef rowNumber As Long) As Boolean

    RepositoryTryResolveLetterRowNumber = False
    rowNumber = 0

    On Error GoTo ResolveError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Letters")

    Dim letterData As Variant
    letterData = RepositoryReadWorksheetMatrix(ws, LetterColumnAddressee, LetterColumnDispatchRegistryDate, LettersTableName)
    If IsEmpty(letterData) Then Exit Function

    Dim startRow As Long
    startRow = RepositoryGetStructuredDataStartRow(ws, LetterColumnAddressee, LetterColumnDispatchRegistryDate, LettersTableName)

    Dim normalizedAddressee As String
    normalizedAddressee = UCase$(Trim$(Addressee))

    Dim normalizedOutgoingNumber As String
    normalizedOutgoingNumber = UCase$(Trim$(OutgoingNumber))

    Dim normalizedOutgoingDate As String
    normalizedOutgoingDate = UCase$(Trim$(OutgoingDate))

    Dim i As Long
    For i = LBound(letterData, 1) To UBound(letterData, 1)
        If UCase$(Trim$(RepositoryMatrixValueOrEmpty(letterData, i, LetterColumnAddressee))) = normalizedAddressee _
           And UCase$(Trim$(RepositoryMatrixValueOrEmpty(letterData, i, LetterColumnOutgoingNumber))) = normalizedOutgoingNumber _
           And UCase$(Trim$(RepositoryMatrixValueOrEmpty(letterData, i, LetterColumnOutgoingDate))) = normalizedOutgoingDate Then
            rowNumber = startRow + i - 1
            RepositoryTryResolveLetterRowNumber = (rowNumber >= FIRST_DATA_ROW)
            Exit Function
        End If
    Next i

    Exit Function

ResolveError:
    RepositoryTryResolveLetterRowNumber = False
    rowNumber = 0
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

    If records Is Nothing Or records.count = 0 Then

        MsgBox t("form.letter_history.msg.no_export_data", "Нет данных для экспорта."), vbExclamation

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

    exportWs.Name = t("form.letter_history.export.sheet_name", "История писем ") & Format$(Date, "dd.mm.yyyy")

    exportWb.Application.Visible = True



    MsgBox t("form.letter_history.msg.export_completed", "Экспорт завершен.") & vbCrLf & _
           t("form.letter_history.msg.records_exported", "Экспортировано записей: ") & records.count, _
           vbInformation, _
           t("form.letter_history.msg.export_title", "Экспорт данных")

    Exit Sub



ExportError:

    MsgBox t("form.letter_history.msg.export_error", "Ошибка экспорта: ") & Err.description, vbCritical

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

        ws.Cells(rowNumber, LetterColumnDocumentSum).value = ""

    ElseIf IsNumeric(sumValue) Then

        ws.Cells(rowNumber, LetterColumnDocumentSum).value = CDbl(sumValue)

    Else

        ws.Cells(rowNumber, LetterColumnDocumentSum).value = sumValue

    End If



    ws.Cells(rowNumber, LetterColumnReturnStatus).value = returnStatus

    Exit Sub



UpdateError:

    Err.Raise Err.Number, "RepositoryUpdateLetterHistoryRow", Err.description

End Sub

Public Sub RepositoryUpdateLetterDispatchTracking( _
    ByVal rowNumber As Long, _
    ByVal packedFlag As String, _
    ByVal batchId As String, _
    ByVal registryNumber As String, _
    ByVal registryDate As String)

    On Error GoTo UpdateError

    If rowNumber < FIRST_DATA_ROW Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Letters")

    ws.Cells(rowNumber, LetterColumnDispatchPackedFlag).value = packedFlag
    ws.Cells(rowNumber, LetterColumnDispatchBatchId).value = batchId
    ws.Cells(rowNumber, LetterColumnDispatchRegistryNumber).value = registryNumber
    ws.Cells(rowNumber, LetterColumnDispatchRegistryDate).value = registryDate
    Exit Sub

UpdateError:
    Err.Raise Err.Number, "RepositoryUpdateLetterDispatchTracking", Err.description
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

    lastRow = ws.Cells(ws.Rows.count, firstColumn).End(xlUp).Row

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



    RepositoryReadWorksheetMatrix = sourceRange.value

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
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPhone) & " " & _
                                       RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnGroup)

End Function



Private Function CreateAddressSearchResultFromMatrix(addressData As Variant, rowIndex As Long, worksheetRowNumber As Long) As Variant

    Dim searchResult(AddressSearchResultGroup) As Variant



    searchResult(AddressSearchResultDisplayText) = BuildAddressSearchDisplayTextFromMatrix(addressData, rowIndex)

    searchResult(AddressSearchResultWorksheetRow) = worksheetRowNumber

    searchResult(AddressSearchResultAddressee) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnAddressee)

    searchResult(AddressSearchResultStreet) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnStreet)

    searchResult(AddressSearchResultCity) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnCity)

    searchResult(AddressSearchResultDistrict) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnDistrict)

    searchResult(AddressSearchResultRegion) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnRegion)

    searchResult(AddressSearchResultPostalCode) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPostalCode)

    searchResult(AddressSearchResultPhone) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPhone)

    searchResult(AddressSearchResultGroup) = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnGroup)



    CreateAddressSearchResultFromMatrix = searchResult

End Function



Private Function BuildAddressSearchDisplayTextFromMatrix(addressData As Variant, rowIndex As Long) As String

    Dim displayText As String

    Dim locationText As String

    Dim groupText As String



    displayText = RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnAddressee)

    locationText = BuildAddressLocationTextFromMatrix(addressData, rowIndex)

    groupText = Trim$(RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnGroup))



    If Len(groupText) > 0 Then

        displayText = displayText & " [" & groupText & "]"

    End If



    If Len(locationText) > 0 Then

        displayText = displayText & " | " & locationText

    End If



    BuildAddressSearchDisplayTextFromMatrix = displayText

End Function



Private Function BuildAddressLocationTextFromMatrix(addressData As Variant, rowIndex As Long) As String

    Dim parts As Collection

    Set parts = New Collection



    AddNonEmptyPart parts, RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnStreet)

    AddNonEmptyPart parts, RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnCity)

    AddNonEmptyPart parts, RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnDistrict)

    AddNonEmptyPart parts, RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnRegion)

    AddNonEmptyPart parts, RepositoryMatrixValueOrEmpty(addressData, rowIndex, AddressColumnPostalCode)



    BuildAddressLocationTextFromMatrix = JoinCollectionParts(parts, ", ")

End Function



Private Function CreateLetterHistoryRecordFromMatrix(letterData As Variant, rowIndex As Long, worksheetRowNumber As Long) As clsLetterHistoryRecord

    Dim record As clsLetterHistoryRecord

    Set record = New clsLetterHistoryRecord



    record.Addressee = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnAddressee)

    record.OutgoingNumber = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnOutgoingNumber)

    record.OutgoingDate = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnOutgoingDate)

    record.attachmentText = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnAttachmentText)

    record.DocumentSum = NormalizeHistorySumCell(letterData(rowIndex, LetterColumnDocumentSum))

    record.returnStatus = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnReturnStatus)

    record.Executor = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnExecutor)

    record.DocumentTypeKey = NormalizeDocumentTypeKey(RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnDocumentType))

    record.DispatchPackedFlag = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnDispatchPackedFlag)

    record.DispatchBatchId = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnDispatchBatchId)

    record.DispatchRegistryNumber = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnDispatchRegistryNumber)

    record.DispatchRegistryDate = RepositoryMatrixValueOrEmpty(letterData, rowIndex, LetterColumnDispatchRegistryDate)

    record.rowNumber = worksheetRowNumber



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

    ws.Cells(rowNumber, AddressColumnAddressee).value = addressArray(AddressIndexAddressee)

    ws.Cells(rowNumber, AddressColumnStreet).value = addressArray(AddressIndexStreet)

    ws.Cells(rowNumber, AddressColumnCity).value = addressArray(AddressIndexCity)

    ws.Cells(rowNumber, AddressColumnDistrict).value = addressArray(AddressIndexDistrict)

    ws.Cells(rowNumber, AddressColumnRegion).value = addressArray(AddressIndexRegion)

    ws.Cells(rowNumber, AddressColumnPostalCode).value = addressArray(AddressIndexPostalCode)

    ws.Cells(rowNumber, AddressColumnPhone).value = FormatPhoneNumber(CStr(addressArray(AddressIndexPhone)))

    ws.Cells(rowNumber, AddressColumnGroup).value = addressArray(AddressIndexGroup)

End Sub



Private Function IsAddressSearchResultArray(addressSearchResult As Variant) As Boolean

    On Error GoTo NotSearchResult



    If Not IsArray(addressSearchResult) Then Exit Function

    If UBound(addressSearchResult) < AddressSearchResultGroup Then Exit Function



    IsAddressSearchResultArray = True

    Exit Function



NotSearchResult:

    IsAddressSearchResultArray = False

End Function



Private Sub AddNonEmptyPart(parts As Collection, textValue As String)

    If Len(Trim$(textValue)) > 0 Then

        parts.Add Trim$(textValue)

    End If

End Sub



Private Function JoinCollectionParts(parts As Collection, delimiter As String) As String

    Dim resultText As String

    Dim i As Long



    For i = 1 To parts.count

        If i > 1 Then resultText = resultText & delimiter

        resultText = resultText & CStr(parts(i))

    Next i



    JoinCollectionParts = resultText

End Function



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

    record.attachmentText = parts(HistoryPartAttachmentText)

    record.DocumentSum = parts(HistoryPartDocumentSum)

    record.returnStatus = parts(HistoryPartReturnStatus)

    record.Executor = parts(HistoryPartExecutor)

    record.DocumentTypeKey = NormalizeDocumentTypeKey(parts(HistoryPartDocumentType))

    record.rowNumber = CLng(parts(HistoryPartRowNumber))



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



    If InStr(1, UCase$(record.attachmentText), searchPattern, vbTextCompare) > 0 Then

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



    If InStr(1, UCase$(record.returnStatus), searchPattern, vbTextCompare) > 0 Then

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

        .Cells(1, LetterColumnAddressee).value = t("form.letter_history.export.header.addressee", "Адресат")

        .Cells(1, LetterColumnOutgoingNumber).value = t("form.letter_history.export.header.outgoing_number", "Исходящий номер")

        .Cells(1, LetterColumnOutgoingDate).value = t("form.letter_history.export.header.outgoing_date", "Дата исходящего")

        .Cells(1, LetterColumnAttachmentText).value = t("form.letter_history.export.header.attachment_name", "Наименование приложения")

        .Cells(1, LetterColumnDocumentSum).value = t("form.letter_history.export.header.document_sum", "Сумма документа")

        .Cells(1, LetterColumnReturnStatus).value = t("form.letter_history.export.header.return_mark", "Отметка о возврате")

        .Cells(1, LetterColumnExecutor).value = t("form.letter_history.export.header.executor_name", "Исполнитель")

        .Cells(1, LetterColumnDocumentType).value = t("form.letter_history.export.header.send_type", "Тип отправки")



        With .Range("A1:H1")

            .Font.Bold = True

            .Interior.ColorIndex = 15

            .Font.ColorIndex = 1

        End With

    End With

End Sub



Private Sub WriteLetterHistoryExportRecords(exportWs As Worksheet, records As Collection)

    Dim i As Long

    For i = 1 To records.count

        WriteLetterHistoryExportRow exportWs, i + 1, records(i)

    Next i

End Sub



Private Sub WriteLetterHistoryExportRow(exportWs As Worksheet, targetRow As Long, letterData As Variant)

    Dim record As clsLetterHistoryRecord

    If Not TryGetLetterHistoryRecord(letterData, record) Then Exit Sub



    exportWs.Cells(targetRow, LetterColumnAddressee).value = record.Addressee

    exportWs.Cells(targetRow, LetterColumnOutgoingNumber).value = record.OutgoingNumber

    exportWs.Cells(targetRow, LetterColumnOutgoingDate).value = record.OutgoingDate

    exportWs.Cells(targetRow, LetterColumnAttachmentText).value = record.attachmentText

    exportWs.Cells(targetRow, LetterColumnDocumentSum).value = record.DocumentSum

    exportWs.Cells(targetRow, LetterColumnReturnStatus).value = record.returnStatus

    exportWs.Cells(targetRow, LetterColumnExecutor).value = record.Executor

    exportWs.Cells(targetRow, LetterColumnDocumentType).value = GetDocumentTypeDisplayLabel(record.DocumentTypeKey)

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

            FormatHistoryDocumentSum = Format$(CDbl(sumText), "#,##0.00") & " руб."

        Else

            FormatHistoryDocumentSum = "-"

        End If

    Else

        FormatHistoryDocumentSum = "-"

    End If

End Function



Private Function BuildHistoryStatusLabel(returnStatus As String) As String

    If RepositoryHasReturnStatusDate(returnStatus) Then

        BuildHistoryStatusLabel = t("history.status.received_label", "Получено ") & returnStatus

    Else

        BuildHistoryStatusLabel = t("history.status.pending_label", "Ожидается ") & returnStatus

    End If

End Function



