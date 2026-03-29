Attribute VB_Name = "ModuleMain"

' ======================================================================
' Module: ModuleMain (main module) - WITH DEBUGGING
' Author: CreateLetter contributors
' Purpose: Core shared logic for validation, data processing, Word generation, and workbook persistence
' Version: 1.7.0 — 29.03.2026
' ======================================================================

Option Explicit

' ======================================================================
'                    SCHEMA CONSTANTS v1.6.21
' ======================================================================
Public Const FIRST_DATA_ROW As Long = 2
Private Const TextTableColumnBody As Long = 1
Private Const TextTableRowOwnForConfirmation As Long = 1
Private Const TextTableRowConfirmedDocuments As Long = 2
Private Const LetterTextsTableName As String = "tblLetterTexts"
Private Const LegacyLetterTextsTableName As String = "Text"
Private Const LegacyLetterTextsTableNameLocalized As String = "Текст"
Public Const AddressesTableName As String = "tblAddresses"
Public Const LettersTableName As String = "tblLetters"
Private Const DocumentTypeKeyConfirmed As String = "confirmed_documents"
Private Const DocumentTypeKeyOwnConfirmation As String = "own_for_confirmation"
Private Const LetterTypeKeyRegular As String = "regular"
Private Const LetterTypeKeyFOU As String = "fou"
Public Const LetterTemplateFileNameRegular As String = "LetterTemplate.docx"
Public Const LetterTemplateFileNameFOU As String = "LetterTemplateFOU.docx"
Public Const TemplatePlaceholderRecipientName As String = "RecipientName"
Public Const TemplatePlaceholderRecipientAddress As String = "RecipientAddress"
Public Const TemplatePlaceholderOutgoingNumber As String = "OutgoingNumber"
Public Const TemplatePlaceholderOutgoingDate As String = "OutgoingDate"
Public Const TemplatePlaceholderExecutorName As String = "ExecutorName"
Public Const TemplatePlaceholderExecutorPhone As String = "ExecutorPhone"
Public Const TemplatePlaceholderLetterText As String = "LetterText"
Public Const TemplatePlaceholderAttachmentsList As String = "AttachmentsList"

Public Enum AddressColumns
    AddressColumnAddressee = 1
    AddressColumnStreet = 2
    AddressColumnCity = 3
    AddressColumnDistrict = 4
    AddressColumnRegion = 5
    AddressColumnPostalCode = 6
    AddressColumnPhone = 7
End Enum

Public Enum LetterColumns
    LetterColumnAddressee = 1
    LetterColumnOutgoingNumber = 2
    LetterColumnOutgoingDate = 3
    LetterColumnAttachmentText = 4
    LetterColumnDocumentSum = 5
    LetterColumnReturnStatus = 6
    LetterColumnExecutor = 7
    LetterColumnDocumentType = 8
End Enum

Public Enum SettingsColumns
    SettingsColumnAttachmentName = 1
    SettingsColumnExecutorName = 3
    SettingsColumnExecutorPhone = 4
End Enum

Public Enum AddressArrayIndexes
    AddressIndexAddressee = 0
    AddressIndexStreet = 1
    AddressIndexCity = 2
    AddressIndexDistrict = 3
    AddressIndexRegion = 4
    AddressIndexPostalCode = 5
    AddressIndexPhone = 6
End Enum

Public Enum AddressListPartIndexes
    AddressPartAddressee = 0
    AddressPartStreet = 1
    AddressPartCity = 2
    AddressPartDistrict = 3
    AddressPartRegion = 4
    AddressPartPostalCode = 5
    AddressPartPhone = 6
    AddressPartRowNumber = 7
End Enum

Public Enum DocumentArrayIndexes
    DocumentIndexName = 0
    DocumentIndexNumber = 1
    DocumentIndexDate = 2
    DocumentIndexCopies = 3
    DocumentIndexSheets = 4
    DocumentIndexSum = 5
End Enum

Public Enum LetterHistoryPartIndexes
    HistoryPartAddressee = 0
    HistoryPartOutgoingNumber = 1
    HistoryPartOutgoingDate = 2
    HistoryPartAttachmentText = 3
    HistoryPartDocumentSum = 4
    HistoryPartReturnStatus = 5
    HistoryPartExecutor = 6
    HistoryPartDocumentType = 7
    HistoryPartRowNumber = 8
End Enum

' ======================================================================
'                    CORE HELPERS v1.6.19
' ======================================================================

Public Sub PopulateDocumentTypeOptions(targetControl As Object)
    On Error Resume Next
    targetControl.Clear
    targetControl.AddItem GetDocumentTypeDisplayLabel(DocumentTypeKeyConfirmed)
    targetControl.AddItem GetDocumentTypeDisplayLabel(DocumentTypeKeyOwnConfirmation)
    targetControl.ListIndex = 0
    On Error GoTo 0
End Sub

Public Sub PopulateLetterTypeOptions(targetControl As Object)
    On Error Resume Next
    targetControl.Clear
    targetControl.AddItem GetLetterTypeDisplayLabel(LetterTypeKeyRegular)
    targetControl.AddItem GetLetterTypeDisplayLabel(LetterTypeKeyFOU)
    targetControl.ListIndex = 0
    On Error GoTo 0
End Sub

Public Function NormalizeDocumentTypeKey(documentType As String) As String
    Dim normalizedText As String
    normalizedText = UCase$(Trim$(documentType))

    If normalizedText = UCase$(DocumentTypeKeyOwnConfirmation) Or _
       normalizedText = UCase$(t("form.letter_creator.option.document_type.own_confirmation", "Own for confirmation")) Or _
       normalizedText = "OWN FOR CONFIRMATION" Then
        NormalizeDocumentTypeKey = DocumentTypeKeyOwnConfirmation
        Exit Function
    End If

    If normalizedText = UCase$(DocumentTypeKeyConfirmed) Or _
       normalizedText = UCase$(t("form.letter_creator.option.document_type.confirmed", "Third-party confirmed documents")) Or _
       normalizedText = "THIRD-PARTY CONFIRMED DOCUMENTS" Then
        NormalizeDocumentTypeKey = DocumentTypeKeyConfirmed
        Exit Function
    End If

    NormalizeDocumentTypeKey = Trim$(documentType)
End Function

Public Function ResolveDocumentTypeStorageValue(documentType As String) As String
    Dim normalizedKey As String
    normalizedKey = NormalizeDocumentTypeKey(documentType)

    If normalizedKey = DocumentTypeKeyOwnConfirmation Or normalizedKey = DocumentTypeKeyConfirmed Then
        ResolveDocumentTypeStorageValue = normalizedKey
    ElseIf Len(Trim$(documentType)) = 0 Then
        ResolveDocumentTypeStorageValue = ""
    Else
        ResolveDocumentTypeStorageValue = DocumentTypeKeyConfirmed
    End If
End Function

Public Function GetDocumentTypeDisplayLabel(documentType As String) As String
    Select Case NormalizeDocumentTypeKey(documentType)
        Case DocumentTypeKeyOwnConfirmation
            GetDocumentTypeDisplayLabel = t("form.letter_creator.option.document_type.own_confirmation", "Own for confirmation")
        Case DocumentTypeKeyConfirmed
            GetDocumentTypeDisplayLabel = t("form.letter_creator.option.document_type.confirmed", "Third-party confirmed documents")
        Case Else
            GetDocumentTypeDisplayLabel = Trim$(documentType)
    End Select
End Function

Public Function NormalizeLetterTypeKey(letterType As String) As String
    Dim normalizedText As String
    normalizedText = UCase$(Trim$(letterType))

    If normalizedText = UCase$(LetterTypeKeyFOU) Or _
       normalizedText = UCase$(t("form.letter_creator.option.letter_type.fou", "FOU (For Official Use)")) Then
        NormalizeLetterTypeKey = LetterTypeKeyFOU
    Else
        NormalizeLetterTypeKey = LetterTypeKeyRegular
    End If
End Function

Public Function GetLetterTypeDisplayLabel(letterType As String) As String
    If NormalizeLetterTypeKey(letterType) = LetterTypeKeyFOU Then
        GetLetterTypeDisplayLabel = t("form.letter_creator.option.letter_type.fou", "FOU (For Official Use)")
    Else
        GetLetterTypeDisplayLabel = t("form.letter_creator.option.letter_type.regular", "Regular")
    End If
End Function

Public Function IsAlternateLetterTypeSelection(letterType As String) As Boolean
    IsAlternateLetterTypeSelection = (NormalizeLetterTypeKey(letterType) = LetterTypeKeyFOU)
End Function


Public Function ValidateRequiredFields(addressee As String, city As String, region As String, postalCode As String, executor As String) As String
    If Len(Trim(addressee)) = 0 Then
        ValidateRequiredFields = t("validation.creator.page.addressee_required", "Fill in the 'Addressee' field.")
        Exit Function
    End If
    
    If Len(Trim(city)) = 0 Then
        ValidateRequiredFields = t("validation.creator.page.city_required", "Fill in the 'City' field. This field is required.")
        Exit Function
    End If
    
    If Len(Trim(region)) = 0 Then
        ValidateRequiredFields = t("validation.creator.page.region_required", "Fill in the 'Region' field. This field is required.")
        Exit Function
    End If
    
    If Len(Trim(postalCode)) = 0 Then
        ValidateRequiredFields = t("validation.creator.page.postal_code_required", "Fill in the 'Postal code' field. This field is required.")
        Exit Function
    End If
    
    ValidateRequiredFields = ""
End Function

Public Function ValidateCreatorPage(pageIndex As Integer, addressee As String, city As String, region As String, postalCode As String, phoneNumber As String, letterNumber As String, letterDateText As String, executor As String, documentsCount As Long, ByRef focusControlName As String) As String
    focusControlName = ""
    ValidateCreatorPage = ""
    
    Select Case pageIndex
        Case 0
            If Len(Trim(addressee)) = 0 Then
                focusControlName = "txtAddressee"
                ValidateCreatorPage = t("validation.creator.page.addressee_required", "Fill in the 'Addressee' field.")
                Exit Function
            End If
            
            If Len(Trim(city)) = 0 Then
                focusControlName = "txtCity"
                ValidateCreatorPage = t("validation.creator.page.city_required", "Fill in the 'City' field. This field is required.")
                Exit Function
            End If
            
            If Len(Trim(region)) = 0 Then
                focusControlName = "txtRegion"
                ValidateCreatorPage = t("validation.creator.page.region_required", "Fill in the 'Region' field. This field is required.")
                Exit Function
            End If
            
            If Len(Trim(postalCode)) = 0 Then
                focusControlName = "txtPostalCode"
                ValidateCreatorPage = t("validation.creator.page.postal_code_required", "Fill in the 'Postal code' field. This field is required.")
                Exit Function
            End If
            
            If Len(Trim(phoneNumber)) > 0 And Not IsPhoneNumberValid(phoneNumber) Then
                focusControlName = "txtAddresseePhone"
                ValidateCreatorPage = t("validation.creator.page.phone_invalid", "Enter a valid addressee phone number.")
                Exit Function
            End If
            
        Case 1
            If Len(Trim(letterNumber)) = 0 Then
                focusControlName = "txtLetterNumber"
                ValidateCreatorPage = t("validation.creator.page.letter_number_required", "Enter the outgoing letter number.")
                Exit Function
            End If
            
            If Len(Trim(letterDateText)) = 0 Then
                focusControlName = "txtLetterDate"
                ValidateCreatorPage = t("validation.creator.page.letter_date_required", "Enter the letter date.")
                Exit Function
            End If
            
            If Len(Trim(executor)) = 0 Then
                focusControlName = "cmbExecutor"
                ValidateCreatorPage = t("validation.creator.page.executor_required", "Select an executor. This field is required.")
                Exit Function
            End If
            
            Dim parsedDate As Date
            If Not TryParseDate(letterDateText, parsedDate) Then
                focusControlName = "txtLetterDate"
                ValidateCreatorPage = t("validation.creator.page.letter_date_invalid", "Invalid letter date format.")
                Exit Function
            End If
            
        Case 2
            If documentsCount = 0 Then
                ValidateCreatorPage = t("validation.creator.page.document_required", "Add at least one attachment document.")
                Exit Function
            End If
    End Select
End Function

Public Function ValidateCreatorSubmission(addressee As String, city As String, region As String, postalCode As String, letterNumber As String, letterDateText As String, executor As String, documentsCount As Long, ByRef focusControlName As String) As String
    focusControlName = ""
    ValidateCreatorSubmission = ""
    
    If Len(Trim(addressee)) = 0 Then
        focusControlName = "txtAddressee"
        ValidateCreatorSubmission = t("validation.creator.submit.addressee_required", "Addressee is not filled in.")
        Exit Function
    End If
    
    If Len(Trim(city)) = 0 Then
        focusControlName = "txtCity"
        ValidateCreatorSubmission = t("validation.creator.submit.city_required", "City is not filled in.")
        Exit Function
    End If
    
    If Len(Trim(region)) = 0 Then
        focusControlName = "txtRegion"
        ValidateCreatorSubmission = t("validation.creator.submit.region_required", "Region is not filled in.")
        Exit Function
    End If
    
    If Len(Trim(postalCode)) = 0 Then
        focusControlName = "txtPostalCode"
        ValidateCreatorSubmission = t("validation.creator.submit.postal_code_required", "Postal code is not filled in.")
        Exit Function
    End If
    
    If Len(Trim(letterNumber)) = 0 Then
        focusControlName = "txtLetterNumber"
        ValidateCreatorSubmission = t("validation.creator.submit.letter_number_required", "Letter number is not filled in.")
        Exit Function
    End If
    
    If Len(Trim(letterDateText)) = 0 Then
        focusControlName = "txtLetterDate"
        ValidateCreatorSubmission = t("validation.creator.submit.letter_date_required", "Letter date is not filled in.")
        Exit Function
    End If
    
    If Len(Trim(executor)) = 0 Then
        focusControlName = "cmbExecutor"
        ValidateCreatorSubmission = t("validation.creator.submit.executor_required", "Executor is not selected.")
        Exit Function
    End If
    
    If documentsCount = 0 Then
        focusControlName = "txtAttachmentSearch"
        ValidateCreatorSubmission = t("validation.creator.submit.document_required", "Add at least one document.")
        Exit Function
    End If
End Function

Public Function FormatPhoneNumber(phoneInput As String) As String
    If Len(Trim(phoneInput)) = 0 Then
        FormatPhoneNumber = ""
        Exit Function
    End If
    
    Dim cleanPhone As String, i As Integer
    For i = 1 To Len(phoneInput)
        If IsNumeric(Mid(phoneInput, i, 1)) Then
            cleanPhone = cleanPhone & Mid(phoneInput, i, 1)
        End If
    Next i
    
    Select Case Len(cleanPhone)
        Case 11
            If Left(cleanPhone, 1) = "8" Or Left(cleanPhone, 1) = "7" Then
                FormatPhoneNumber = Left(cleanPhone, 1) & "-" & _
                                  Mid(cleanPhone, 2, 3) & "-" & _
                                  Mid(cleanPhone, 5, 3) & "-" & _
                                  Mid(cleanPhone, 8, 2) & "-" & _
                                  Mid(cleanPhone, 10, 2)
            Else
                FormatPhoneNumber = cleanPhone
            End If
            
        Case 10
            FormatPhoneNumber = "8-" & Left(cleanPhone, 3) & "-" & _
                              Mid(cleanPhone, 4, 3) & "-" & _
                              Mid(cleanPhone, 7, 2) & "-" & _
                              Mid(cleanPhone, 9, 2)
                              
        Case 7
            FormatPhoneNumber = Left(cleanPhone, 3) & "-" & _
                              Mid(cleanPhone, 4, 2) & "-" & _
                              Mid(cleanPhone, 6, 2)
                              
        Case Else
            FormatPhoneNumber = phoneInput
    End Select
End Function

Public Function IsPhoneNumberValid(phoneNumber As String) As Boolean
    Dim cleanPhone As String, i As Integer
    
    For i = 1 To Len(phoneNumber)
        If IsNumeric(Mid(phoneNumber, i, 1)) Then
            cleanPhone = cleanPhone & Mid(phoneNumber, i, 1)
        End If
    Next i
    
    IsPhoneNumberValid = (Len(cleanPhone) >= 7 And Len(cleanPhone) <= 11)
End Function

' ======================================================================
'                    DOCUMENT FUNCTIONS
' ======================================================================
Public Function CreateDocumentArray(docName As String, docNumber As String, docDate As String, docCopies As String, docSheets As String) As Variant
    Dim docArray(DocumentIndexSheets) As String
    docArray(DocumentIndexName) = Trim(docName)
    docArray(DocumentIndexNumber) = Trim(docNumber)
    docArray(DocumentIndexDate) = Trim(docDate)
    docArray(DocumentIndexCopies) = Trim(docCopies)
    docArray(DocumentIndexSheets) = Trim(docSheets)
    
    CreateDocumentArray = docArray
End Function

Public Function CreateDocumentArrayWithSum(docName As String, docNumber As String, docDate As String, docCopies As String, docSheets As String, docSum As String) As Variant
    Dim docArray(DocumentIndexSum) As String
    docArray(DocumentIndexName) = Trim(docName)
    docArray(DocumentIndexNumber) = Trim(docNumber)
    docArray(DocumentIndexDate) = Trim(docDate)
    docArray(DocumentIndexCopies) = Trim(docCopies)
    docArray(DocumentIndexSheets) = Trim(docSheets)
    docArray(DocumentIndexSum) = Trim(docSum)
    
    CreateDocumentArrayWithSum = docArray
End Function

Public Function FormatDocumentName(docArray As Variant) As String
    If Not IsArray(docArray) Then
        FormatDocumentName = t("core.runtime.error.invalid_data_format", "Error: invalid data format")
        Exit Function
    End If
    
    Dim result As String
    result = docArray(DocumentIndexName)
    
    result = result & " No."
    If Len(Trim(docArray(DocumentIndexNumber))) > 0 Then
        result = result & docArray(DocumentIndexNumber)
    Else
        result = result & "    "
    End If
    
    result = result & " dated "
    If Len(Trim(docArray(DocumentIndexDate))) > 0 Then
        result = result & docArray(DocumentIndexDate)
    Else
        result = result & "        "
    End If
    
    result = result & " ("
    
    If Len(Trim(docArray(DocumentIndexCopies))) > 0 Then
        result = result & docArray(DocumentIndexCopies) & " copies"
    Else
        result = result & "  copies"
    End If
    
    result = result & ", "
    If Len(Trim(docArray(DocumentIndexSheets))) > 0 Then
        result = result & docArray(DocumentIndexSheets) & " sheets"
    Else
        result = result & "   sheets"
    End If
    
    result = result & ")"
    
    FormatDocumentName = result
End Function

Private Function TryGetWorksheetTable(ws As Worksheet, tableName As String, ByRef targetTable As ListObject) As Boolean
    On Error GoTo TableMissing
    
    Set targetTable = ws.ListObjects(tableName)
    TryGetWorksheetTable = Not targetTable Is Nothing
    Exit Function
    
TableMissing:
    Set targetTable = Nothing
    TryGetWorksheetTable = False
End Function

Public Function GetAddressesTable() As ListObject
    Dim ws As Worksheet
    Dim targetTable As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Addresses")
    If TryGetWorksheetTable(ws, AddressesTableName, targetTable) Then
        Set GetAddressesTable = targetTable
    End If
End Function

Public Function GetLettersTable() As ListObject
    Dim ws As Worksheet
    Dim targetTable As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Letters")
    If TryGetWorksheetTable(ws, LettersTableName, targetTable) Then
        Set GetLettersTable = targetTable
    End If
End Function

Public Function GetStructuredTableReadiness() As String
    Dim readiness As String
    readiness = "Addresses="
    readiness = readiness & IIf(GetAddressesTable() Is Nothing, "missing", "present")
    readiness = readiness & ";Letters="
    readiness = readiness & IIf(GetLettersTable() Is Nothing, "missing", "present")
    GetStructuredTableReadiness = readiness
End Function

Private Function GetStructuredDataRange(ws As Worksheet, firstColumn As Long, lastColumn As Long, Optional preferredTableName As String = "") As Range
    If Len(preferredTableName) > 0 Then
        Dim preferredTable As ListObject
        If TryGetWorksheetTable(ws, preferredTableName, preferredTable) Then
            If Not preferredTable.DataBodyRange Is Nothing Then
                Set GetStructuredDataRange = preferredTable.DataBodyRange
                Exit Function
            End If
        End If
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, firstColumn).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Function
    
    Set GetStructuredDataRange = ws.Range(ws.Cells(FIRST_DATA_ROW, firstColumn), ws.Cells(lastRow, lastColumn))
End Function

Private Function ReadWorksheetMatrix(ws As Worksheet, firstColumn As Long, lastColumn As Long, Optional preferredTableName As String = "") As Variant
    Dim sourceRange As Range
    Set sourceRange = GetStructuredDataRange(ws, firstColumn, lastColumn, preferredTableName)
    
    If sourceRange Is Nothing Then
        ReadWorksheetMatrix = Empty
        Exit Function
    End If
    
    ReadWorksheetMatrix = sourceRange.value
End Function

Private Function GetStructuredDataStartRow(ws As Worksheet, firstColumn As Long, lastColumn As Long, Optional preferredTableName As String = "") As Long
    Dim sourceRange As Range
    Set sourceRange = GetStructuredDataRange(ws, firstColumn, lastColumn, preferredTableName)
    
    If sourceRange Is Nothing Then
        GetStructuredDataStartRow = FIRST_DATA_ROW
    Else
        GetStructuredDataStartRow = sourceRange.Row
    End If
End Function

Private Function MatrixValueOrEmpty(dataMatrix As Variant, rowIndex As Long, columnIndex As Long) As String
    If IsArray(dataMatrix) Then
        MatrixValueOrEmpty = CStr(dataMatrix(rowIndex, columnIndex))
    Else
        MatrixValueOrEmpty = ""
    End If
End Function

Private Function BuildAddressSearchLineFromMatrix(addressData As Variant, rowIndex As Long) As String
    BuildAddressSearchLineFromMatrix = MatrixValueOrEmpty(addressData, rowIndex, AddressColumnAddressee) & " " & _
                                       MatrixValueOrEmpty(addressData, rowIndex, AddressColumnStreet) & " " & _
                                       MatrixValueOrEmpty(addressData, rowIndex, AddressColumnCity) & " " & _
                                       MatrixValueOrEmpty(addressData, rowIndex, AddressColumnDistrict) & " " & _
                                       MatrixValueOrEmpty(addressData, rowIndex, AddressColumnRegion) & " " & _
                                       MatrixValueOrEmpty(addressData, rowIndex, AddressColumnPostalCode) & " " & _
                                       MatrixValueOrEmpty(addressData, rowIndex, AddressColumnPhone)
End Function

Private Function BuildAddressListItemFromMatrix(addressData As Variant, rowIndex As Long, worksheetRowNumber As Long) As String
    BuildAddressListItemFromMatrix = MatrixValueOrEmpty(addressData, rowIndex, AddressColumnAddressee) & " | " & _
                                     MatrixValueOrEmpty(addressData, rowIndex, AddressColumnStreet) & " | " & _
                                     MatrixValueOrEmpty(addressData, rowIndex, AddressColumnCity) & " | " & _
                                     MatrixValueOrEmpty(addressData, rowIndex, AddressColumnDistrict) & " | " & _
                                     MatrixValueOrEmpty(addressData, rowIndex, AddressColumnRegion) & " | " & _
                                     MatrixValueOrEmpty(addressData, rowIndex, AddressColumnPostalCode) & " | " & _
                                     MatrixValueOrEmpty(addressData, rowIndex, AddressColumnPhone) & " | " & worksheetRowNumber
End Function

Private Function BuildLetterHistoryRecordFromMatrix(letterData As Variant, rowIndex As Long, worksheetRowNumber As Long) As String
    BuildLetterHistoryRecordFromMatrix = MatrixValueOrEmpty(letterData, rowIndex, LetterColumnAddressee) & "|" & _
                                         MatrixValueOrEmpty(letterData, rowIndex, LetterColumnOutgoingNumber) & "|" & _
                                         MatrixValueOrEmpty(letterData, rowIndex, LetterColumnOutgoingDate) & "|" & _
                                         MatrixValueOrEmpty(letterData, rowIndex, LetterColumnAttachmentText) & "|" & _
                                         NormalizeHistorySumCell(letterData(rowIndex, LetterColumnDocumentSum)) & "|" & _
                                         MatrixValueOrEmpty(letterData, rowIndex, LetterColumnReturnStatus) & "|" & _
                                         MatrixValueOrEmpty(letterData, rowIndex, LetterColumnExecutor) & "|" & _
                                         MatrixValueOrEmpty(letterData, rowIndex, LetterColumnDocumentType) & "|" & _
                                         CStr(worksheetRowNumber)
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

' ======================================================================
'                    SEARCH AND DATA FUNCTIONS
' ======================================================================
Public Function SearchAddresses(searchTerm As String) As Collection
    Set SearchAddresses = RepositorySearchAddresses(searchTerm)
End Function

Public Function TryParseAddressListItem(addressItem As String, ByRef addressParts As Variant, ByRef rowNumber As Long, ByRef errorMessage As String) As Boolean
    errorMessage = ""
    rowNumber = 0
    TryParseAddressListItem = False
    
    If InStr(addressItem, " | ") = 0 Then
        errorMessage = t("validation.address.record_invalid", "Invalid address record format.")
        Exit Function
    End If
    
    addressParts = Split(addressItem, " | ")
    If UBound(addressParts) < 7 Then
        errorMessage = t("validation.address.record_incomplete", "Address data is incomplete.")
        Exit Function
    End If
    
    If Not IsNumeric(addressParts(AddressPartRowNumber)) Then
        errorMessage = t("validation.address.row_invalid", "Address row reference is invalid.")
        Exit Function
    End If
    
    rowNumber = CLng(addressParts(AddressPartRowNumber))
    TryParseAddressListItem = True
End Function

Public Function ValidateAddressCreateRequest(addressee As String, isDuplicate As Boolean) As String
    If Len(Trim(addressee)) = 0 Then
        ValidateAddressCreateRequest = t("validation.address.create.addressee_required", "Enter the addressee name.")
        Exit Function
    End If
    
    If isDuplicate Then
        ValidateAddressCreateRequest = t("validation.address.create.duplicate", "This address already exists.")
        Exit Function
    End If
    
    ValidateAddressCreateRequest = ""
End Function

Public Function ValidateAddressEditRequest(selectedRow As Long, isDuplicate As Boolean) As String
    If selectedRow <= 1 Then
        ValidateAddressEditRequest = t("validation.address.edit.selection_required", "Select an address to edit.")
        Exit Function
    End If
    
    If isDuplicate Then
        ValidateAddressEditRequest = t("validation.address.edit.duplicate", "An address with the same data already exists.")
        Exit Function
    End If
    
    ValidateAddressEditRequest = ""
End Function

Public Function ValidateAddressDeleteRequest(selectedRow As Long) As String
    If selectedRow = 0 Then
        ValidateAddressDeleteRequest = t("validation.address.delete.selection_required", "Select an address to delete.")
    Else
        ValidateAddressDeleteRequest = ""
    End If
End Function

Public Function IsAddressReadyForAutoUpdate(city As String, region As String, postalCode As String) As Boolean
    IsAddressReadyForAutoUpdate = (Len(Trim(city)) > 0 And Len(Trim(region)) > 0 And Len(Trim(postalCode)) > 0)
End Function

Public Function LoadLetterHistoryData() As Collection
    Set LoadLetterHistoryData = RepositoryLoadLetterHistoryData()
End Function

Public Function FilterLetterHistoryRecords(allLettersData As Collection, searchText As String) As Collection
    Set FilterLetterHistoryRecords = RepositoryFilterLetterHistoryRecords(allLettersData, searchText)
End Function

Public Function FormatLetterHistoryDisplay(letterData As Variant) As String
    FormatLetterHistoryDisplay = RepositoryFormatLetterHistoryDisplay(letterData)
End Function

Public Function TryParseLetterHistoryRecord(letterData As Variant, ByRef parts As Variant) As Boolean
    TryParseLetterHistoryRecord = RepositoryTryParseLetterHistoryRecord(letterData, parts)
End Function

Public Function BuildLetterReturnStatus(isReceived As Boolean, returnDateText As String) As String
    BuildLetterReturnStatus = RepositoryBuildLetterReturnStatus(isReceived, returnDateText)
End Function

Public Function GetLetterHistorySearchHintsText() As String
    GetLetterHistorySearchHintsText = RepositoryGetLetterHistorySearchHintsText()
End Function

Public Sub ExportLetterHistoryRecords(records As Collection)
    RepositoryExportLetterHistoryRecords records
End Sub

Public Function HasReturnStatusDate(returnStatus As String) As Boolean
    HasReturnStatusDate = RepositoryHasReturnStatusDate(returnStatus)
End Function

Public Function ExtractReturnStatusDate(returnStatus As String) As String
    ExtractReturnStatusDate = RepositoryExtractReturnStatusDate(returnStatus)
End Function

Public Sub UpdateLetterHistoryRow(rowNumber As Long, sumValue As String, returnStatus As String)
    RepositoryUpdateLetterHistoryRow rowNumber, sumValue, returnStatus
End Sub

' History record parsing/display/export moved to ModuleRepository in v1.7.0.

Public Function SearchAttachments(searchTerm As String) As Collection
    Set SearchAttachments = RepositorySearchAttachments(searchTerm)
End Function

' ======================================================================
'                    EXECUTOR FUNCTIONS
' ======================================================================
Public Function GetExecutorsList() As Collection
    Set GetExecutorsList = RepositoryGetExecutorsList()
End Function

Public Function GetCurrentUserFIO() As String
    On Error Resume Next
    GetCurrentUserFIO = Environ("USERNAME")
    If GetCurrentUserFIO = "" Then GetCurrentUserFIO = t("common.unknown_user", "Unknown user")
    On Error GoTo 0
End Function

Public Function GetExecutorPhone(executorFIO As String) As String
    GetExecutorPhone = RepositoryGetExecutorPhone(executorFIO)
End Function

Private Sub WriteAddressRow(ws As Worksheet, rowNumber As Long, addressArray As Variant)
    ws.Cells(rowNumber, AddressColumnAddressee).value = addressArray(AddressIndexAddressee)
    ws.Cells(rowNumber, AddressColumnStreet).value = addressArray(AddressIndexStreet)
    ws.Cells(rowNumber, AddressColumnCity).value = addressArray(AddressIndexCity)
    ws.Cells(rowNumber, AddressColumnDistrict).value = addressArray(AddressIndexDistrict)
    ws.Cells(rowNumber, AddressColumnRegion).value = addressArray(AddressIndexRegion)
    ws.Cells(rowNumber, AddressColumnPostalCode).value = addressArray(AddressIndexPostalCode)
    ws.Cells(rowNumber, AddressColumnPhone).value = FormatPhoneNumber(CStr(addressArray(AddressIndexPhone)))
End Sub

Private Function AddressColumnFromIndex(addressIndex As AddressArrayIndexes) As AddressColumns
    Select Case addressIndex
        Case AddressIndexAddressee: AddressColumnFromIndex = AddressColumnAddressee
        Case AddressIndexStreet: AddressColumnFromIndex = AddressColumnStreet
        Case AddressIndexCity: AddressColumnFromIndex = AddressColumnCity
        Case AddressIndexDistrict: AddressColumnFromIndex = AddressColumnDistrict
        Case AddressIndexRegion: AddressColumnFromIndex = AddressColumnRegion
        Case AddressIndexPostalCode: AddressColumnFromIndex = AddressColumnPostalCode
        Case Else: AddressColumnFromIndex = AddressColumnPhone
    End Select
End Function

' ======================================================================
'                    DATA SAVING FUNCTIONS
' ======================================================================
Public Sub SaveNewAddress(addressArray As Variant)
    RepositorySaveNewAddress addressArray
End Sub

Public Sub UpdateExistingAddress(rowNumber As Long, addressArray As Variant)
    RepositoryUpdateExistingAddress rowNumber, addressArray
End Sub

Public Sub DeleteExistingAddress(rowNumber As Long)
    RepositoryDeleteExistingAddress rowNumber
End Sub

Public Function IsAddressDuplicate(addressArray As Variant, Optional excludeRow As Long = 0) As Boolean
    IsAddressDuplicate = RepositoryIsAddressDuplicate(addressArray, excludeRow)
End Function

' ======================================================================
'                    DEBUGGING FUNCTIONS
' ======================================================================

Public Sub SaveLetterInfoWithSum(addressee As String, letterNumber As String, letterDate As Date, documents As Collection, executor As String, documentType As String)
    ' === DEBUG START ===
    Debug.Print "=== DEBUG SaveLetterInfoWithSum START ==="
    Debug.Print "Addressee: " & addressee
    Debug.Print "LetterNumber: " & letterNumber
    Debug.Print "LetterDate: " & letterDate
    Debug.Print "Executor: " & executor
    Debug.Print "DocumentType: " & documentType
    Debug.Print "Documents count: " & documents.count
    
    Dim i As Long
    For i = 1 To documents.count
        Dim docArray As Variant
        docArray = documents(i)
        
        Debug.Print "Document #" & i & ": UBound=" & UBound(docArray) & " LBound=" & LBound(docArray)
        
        Dim j As Long
        For j = LBound(docArray) To UBound(docArray)
            Debug.Print "  Element " & j & ": '" & CStr(docArray(j)) & "'"
        Next j
    Next i
    Debug.Print "=== DEBUG SaveLetterInfoWithSum INITIAL END ==="
    ' === DEBUG END ===
    
    On Error GoTo SaveLetterError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Letters")
    
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.count, LetterColumnAddressee).End(xlUp).Row + 1
    
    ' DEBUG: Writing basic data
    Debug.Print "=== BEFORE writing basic data ==="
    ws.Cells(newRow, LetterColumnAddressee).value = addressee
    ws.Cells(newRow, LetterColumnOutgoingNumber).value = letterNumber
    ws.Cells(newRow, LetterColumnOutgoingDate).value = letterDate
    Debug.Print "=== AFTER writing basic data ==="
    
    ' DEBUG: Formatting attachments
    Debug.Print "=== BEFORE FormatAttachmentsListCompactWithSum ==="
    Dim attachmentText As String
    attachmentText = FormatAttachmentsListCompactWithSum(documents)
    Debug.Print "=== AFTER FormatAttachmentsListCompactWithSum ==="
    Debug.Print "Result length: " & Len(attachmentText)
    Debug.Print "Result preview: " & Left(attachmentText, 200)
    
    ' DEBUG: Writing to Excel
    Debug.Print "=== BEFORE writing to Excel cell (4) ==="
    ws.Cells(newRow, LetterColumnAttachmentText).value = attachmentText
    Debug.Print "=== AFTER writing to Excel cell (4) ==="
    
    ' DEBUG: Sum calculation
    Debug.Print "=== BEFORE CalculateTotalDocumentsSum ==="
    Dim totalSum As Double
    totalSum = CalculateTotalDocumentsSum(documents)
    Debug.Print "=== AFTER CalculateTotalDocumentsSum, result: " & totalSum
    
    ' DEBUG: Writing sum
    Debug.Print "=== BEFORE writing sum to cell (5) ==="
    If totalSum > 0 Then
        ws.Cells(newRow, LetterColumnDocumentSum).value = totalSum
        Debug.Print "Written totalSum: " & totalSum
    Else
        ws.Cells(newRow, LetterColumnDocumentSum).value = ""
        Debug.Print "Written empty sum"
    End If
    Debug.Print "=== AFTER writing sum to cell (5) ==="
    
    ' DEBUG: Writing remaining data
    Debug.Print "=== BEFORE writing remaining cells ==="
    ws.Cells(newRow, LetterColumnReturnStatus).value = ""
    ws.Cells(newRow, LetterColumnExecutor).value = executor
    ws.Cells(newRow, LetterColumnDocumentType).value = documentType
    Debug.Print "=== AFTER writing remaining cells ==="
    
    Debug.Print "=== DEBUG SaveLetterInfoWithSum SUCCESS END ==="
    
    Exit Sub
    
SaveLetterError:
    Debug.Print "=== ERROR in SaveLetterInfoWithSum ==="
    Debug.Print "Error Number: " & Err.number
    Debug.Print "Error Description: " & Err.description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "==========================="
    MsgBox t("core.letter.error.save_info", "Error saving letter info: ") & Err.description, vbCritical
End Sub

Public Function FormatAttachmentsListCompactWithSum(documentsList As Collection) As String
    Debug.Print "=== DEBUG FormatAttachmentsListCompactWithSum START ==="
    
    If documentsList Is Nothing Or documentsList.count = 0 Then
        FormatAttachmentsListCompactWithSum = t("core.attachments.not_specified", "Documents not specified")
        Debug.Print "=== DEBUG FormatAttachmentsListCompactWithSum END (empty) ==="
        Exit Function
    End If
    
    Dim result As String
    Dim i As Long
    
    For i = 1 To documentsList.count
        Debug.Print "Processing document " & i & " of " & documentsList.count
        
        If i > 1 Then result = result & "; "
        
        Dim docArray As Variant
        docArray = documentsList(i)
        
        If IsArray(docArray) And UBound(docArray) >= DocumentIndexSum Then
            Debug.Print "  Calling FormatDocumentNameWithSum"
            Dim docResult As String
            docResult = FormatDocumentNameWithSum(docArray)
            result = result & docResult & ";"
            Debug.Print "  Result so far: " & result
        Else
            Debug.Print "  Calling FormatDocumentName"
            result = result & FormatDocumentName(docArray) & ";"
            Debug.Print "  Result so far: " & result
        End If
    Next i
    
    Debug.Print "Final result length: " & Len(result)
    Debug.Print "Final result: " & result
    Debug.Print "=== DEBUG FormatAttachmentsListCompactWithSum END ==="
    
    FormatAttachmentsListCompactWithSum = result
End Function

Public Function FormatDocumentNameWithSum(docArray As Variant) As String
    Debug.Print "=== DEBUG FormatDocumentNameWithSum START ==="
    Debug.Print "IsArray: " & IsArray(docArray)
    
    If Not IsArray(docArray) Then
        FormatDocumentNameWithSum = t("core.runtime.error.invalid_data_format", "Error: invalid data format")
        Debug.Print "ERROR: Not array"
        Debug.Print "=== DEBUG FormatDocumentNameWithSum END ==="
        Exit Function
    End If
    
    Debug.Print "Array UBound: " & UBound(docArray) & " LBound: " & LBound(docArray)
    
    Dim j As Long
    For j = LBound(docArray) To UBound(docArray)
        Debug.Print "  Element " & j & ": '" & CStr(docArray(j)) & "'"
    Next j
    
    Dim result As String
    result = docArray(DocumentIndexName)
    
    result = result & " No."
    If Len(Trim(docArray(DocumentIndexNumber))) > 0 Then
        result = result & docArray(DocumentIndexNumber)
    Else
        result = result & "    "
    End If
    
    result = result & " dated "
    If Len(Trim(docArray(DocumentIndexDate))) > 0 Then
        result = result & docArray(DocumentIndexDate)
    Else
        result = result & "        "
    End If
    
    ' FIXED SUM CHECK
    If UBound(docArray) >= DocumentIndexSum And Len(Trim(docArray(DocumentIndexSum))) > 0 Then
        Debug.Print "Processing sum: '" & docArray(DocumentIndexSum) & "'"
        If IsNumeric(docArray(DocumentIndexSum)) Then
            Dim sumText As String
            sumText = CStr(CLng(CDbl(docArray(DocumentIndexSum))))
            result = result & " for the amount of " & sumText & " rub."
            Debug.Print "Sum formatted as: " & sumText
        Else
            result = result & " (" & docArray(DocumentIndexSum) & ")"
            Debug.Print "Sum as text: " & docArray(DocumentIndexSum)
        End If
    Else
        Debug.Print "No sum found or empty sum"
    End If
    
    result = result & " ("
    
    If Len(Trim(docArray(DocumentIndexCopies))) > 0 Then
        result = result & docArray(DocumentIndexCopies) & " copies"
    Else
        result = result & "  copies"
    End If
    
    result = result & ", "
    If Len(Trim(docArray(DocumentIndexSheets))) > 0 Then
        result = result & docArray(DocumentIndexSheets) & " sheets"
    Else
        result = result & "   sheets"
    End If
    
    result = result & ")"
    
    Debug.Print "Final document result: " & result
    Debug.Print "=== DEBUG FormatDocumentNameWithSum END ==="
    
    FormatDocumentNameWithSum = result
End Function

Public Function FormatAttachmentsListForWordWithSum(documentsList As Collection) As Collection
    Set FormatAttachmentsListForWordWithSum = New Collection
    
    If documentsList Is Nothing Or documentsList.count = 0 Then
        FormatAttachmentsListForWordWithSum.Add t("core.attachments.not_specified_word", "documents not specified;")
        Exit Function
    End If
    
    Dim currentFragment As String
    Dim i As Long
    Dim docText As String
    
    For i = 1 To documentsList.count
        docText = i & "). " & FormatDocumentNameWithSum(documentsList(i)) & ";"
        
        If Len(currentFragment & vbCrLf & docText) > 180 Then
            If Len(currentFragment) > 0 Then
                FormatAttachmentsListForWordWithSum.Add currentFragment
                currentFragment = ""
            End If
        End If
        
        If Len(currentFragment) > 0 Then
            currentFragment = currentFragment & vbCrLf
        End If
        
        currentFragment = currentFragment & docText
    Next i
    
    If Len(currentFragment) > 0 Then
        FormatAttachmentsListForWordWithSum.Add currentFragment
    End If
End Function

Public Function BuildSummaryAttachmentsText(documentsList As Collection) As String
    If documentsList Is Nothing Or documentsList.count = 0 Then
        BuildSummaryAttachmentsText = ""
        Exit Function
    End If
    
    Dim attachmentText As String
    Dim i As Long
    
    For i = 1 To documentsList.count
        If i > 1 Then attachmentText = attachmentText & vbCrLf
        attachmentText = attachmentText & i & ". " & FormatDocumentNameWithSum(documentsList(i)) & ";"
    Next i
    
    BuildSummaryAttachmentsText = attachmentText
End Function

Public Function BuildCreatorProgressCaption(currentStep As Long, totalPages As Long) As String
    BuildCreatorProgressCaption = t("form.letter_creator.progress.page", "Step") & " " & currentStep & " " & t("common.of", "of") & " " & totalPages
End Function

Public Function BuildCreatorSelectedDocumentsCaption(documentCount As Long) As String
    BuildCreatorSelectedDocumentsCaption = t("form.letter_creator.attachments_count", "Selected documents:") & " " & documentCount
End Function

Public Function GetDocumentActionsMenuPrompt() As String
    GetDocumentActionsMenuPrompt = t("form.letter_creator.menu.document_actions_prompt", "Select action:") & vbCrLf & _
                                   t("form.letter_creator.menu.document_action.edit", "1 - Edit details") & vbCrLf & _
                                   t("form.letter_creator.menu.document_action.duplicate", "2 - Duplicate document") & vbCrLf & _
                                   t("form.letter_creator.menu.document_action.remove", "3 - Remove from list") & vbCrLf & _
                                   t("form.letter_creator.menu.document_action.move_up", "4 - Move up") & vbCrLf & _
                                   t("form.letter_creator.menu.document_action.move_down", "5 - Move down")
End Function

Public Function GetDocumentActionsMenuTitle() As String
    GetDocumentActionsMenuTitle = t("form.letter_creator.menu.document_actions_title", "Document actions")
End Function

Public Function GetDocumentDisplayItems(documentsList As Collection) As Collection
    Set GetDocumentDisplayItems = New Collection
    
    If documentsList Is Nothing Then Exit Function
    
    Dim i As Long
    For i = 1 To documentsList.count
        GetDocumentDisplayItems.Add FormatDocumentNameWithSum(documentsList(i))
    Next i
End Function

Public Function DuplicateDocumentArray(sourceItem As Variant) As Variant
    Dim sourceName As String
    Dim sourceDate As String
    Dim sourceCopies As String
    Dim sourceSheets As String
    Dim sourceSum As String
    
    sourceName = ""
    sourceDate = ""
    sourceCopies = ""
    sourceSheets = ""
    sourceSum = ""
    
    If IsArray(sourceItem) Then
        If UBound(sourceItem) >= DocumentIndexSheets Then
            sourceName = CStr(sourceItem(DocumentIndexName))
            sourceDate = CStr(sourceItem(DocumentIndexDate))
            sourceCopies = CStr(sourceItem(DocumentIndexCopies))
            sourceSheets = CStr(sourceItem(DocumentIndexSheets))
        End If
        
        If UBound(sourceItem) >= DocumentIndexSum Then
            sourceSum = CStr(sourceItem(DocumentIndexSum))
        End If
    End If
    
    DuplicateDocumentArray = CreateDocumentArrayWithSum(sourceName, "", sourceDate, sourceCopies, sourceSheets, sourceSum)
End Function

Public Sub MoveDocumentCollectionItemUp(documentsList As Collection, oneBasedIndex As Long)
    If documentsList Is Nothing Then Exit Sub
    If oneBasedIndex <= 1 Or oneBasedIndex > documentsList.count Then Exit Sub
    
    Dim tempDoc As Variant
    tempDoc = documentsList(oneBasedIndex - 1)
    documentsList.Remove oneBasedIndex - 1
    documentsList.Add tempDoc, , oneBasedIndex - 1
End Sub

Public Sub MoveDocumentCollectionItemDown(documentsList As Collection, oneBasedIndex As Long)
    If documentsList Is Nothing Then Exit Sub
    If oneBasedIndex < 1 Or oneBasedIndex >= documentsList.count Then Exit Sub
    
    Dim tempDoc As Variant
    tempDoc = documentsList(oneBasedIndex + 1)
    documentsList.Remove oneBasedIndex + 1
    documentsList.Add tempDoc, , oneBasedIndex
End Sub

Public Function GetSharedWordApplication() As Object
    Set GetSharedWordApplication = AcquireWordApplication()
End Function

Public Sub ResetSharedWordApplication()
    ResetStaleWordApplication
End Sub

Public Sub ReleaseSharedWordApplication(Optional closeDocuments As Boolean = False)
    ReleaseWordApplication closeDocuments
End Sub

Public Function GetSharedWordApplicationState() As String
    GetSharedWordApplicationState = GetWordApplicationState()
End Function

Public Function WarmUpSharedWordApplication() As Boolean
    WarmUpSharedWordApplication = WarmUpWordApplication()
End Function

Private Function TryGetLoadedUserForm(formName As String, ByRef loadedForm As Object) As Boolean
    On Error GoTo LookupFailed
    
    Set loadedForm = VBA.UserForms(formName)
    TryGetLoadedUserForm = Not loadedForm Is Nothing
    Exit Function
    
LookupFailed:
    Set loadedForm = Nothing
    TryGetLoadedUserForm = False
End Function

Public Sub CreateLetterDocument(addressee As String, addressArray As Variant, letterNumber As String, letterDateRaw As String, executor As String, documentType As String, useAlternateTemplate As Boolean, documentsList As Collection)
    WordInteropCreateLetterDocument addressee, addressArray, letterNumber, letterDateRaw, executor, documentType, useAlternateTemplate, documentsList
End Sub

Public Sub FillWordTemplateData(wordDoc As Object, addresseeText As String, addressArray As Variant, numberText As String, rawDateText As String, executorText As String, documentType As String, documentsList As Collection)
    WordInteropFillWordTemplateData wordDoc, addresseeText, addressArray, numberText, rawDateText, executorText, documentType, documentsList
End Sub

Public Sub CreateLetterDocumentFromScratch(wordDoc As Object, addresseeText As String, addressArray As Variant, numberText As String, rawDateText As String, executorText As String, documentType As String, documentsList As Collection)
    WordInteropCreateLetterDocumentFromScratch wordDoc, addresseeText, addressArray, numberText, rawDateText, executorText, documentType, documentsList
End Sub

Public Sub ReplaceAttachmentsInTemplateWithFontAndSum(wordDoc As Object, documentsList As Collection, fontSize As Integer)
    WordInteropReplaceAttachmentsInTemplateWithFontAndSum wordDoc, documentsList, fontSize
End Sub

Public Sub AppendAttachmentsToDocumentWithFontAndSum(wordDoc As Object, documentsList As Collection, fontSize As Integer)
    WordInteropAppendAttachmentsToDocumentWithFontAndSum wordDoc, documentsList, fontSize
End Sub

Public Function CalculateTotalDocumentsSum(documents As Collection) As Double
    Debug.Print "=== DEBUG CalculateTotalDocumentsSum START ==="
    CalculateTotalDocumentsSum = 0
    
    If documents Is Nothing Or documents.count = 0 Then
        Debug.Print "=== DEBUG CalculateTotalDocumentsSum END (empty) ==="
        Exit Function
    End If
    
    Dim documentsWithSum As Integer
    Dim totalSum As Double
    documentsWithSum = 0
    totalSum = 0
    
    Dim i As Long
    For i = 1 To documents.count
        Dim docArray As Variant
        docArray = documents(i)
        
        Debug.Print "Checking document " & i & " for sum calculation"
        
        If IsArray(docArray) And UBound(docArray) >= DocumentIndexSum Then
            Dim docSum As String
            docSum = Trim(CStr(docArray(DocumentIndexSum)))
            
            Debug.Print "  Document " & i & " sum: '" & docSum & "'"
            Debug.Print "  IsNumeric: " & IsNumeric(docSum)
            Debug.Print "  Len > 0: " & (Len(docSum) > 0)
            
            ' FIXED: Breaking down condition into nested IFs to avoid premature CDbl call
            If Len(docSum) > 0 Then
                If IsNumeric(docSum) Then
                    Dim sumValue As Double
                    sumValue = CDbl(docSum)
                    Debug.Print "  Converted sum: " & sumValue
                    
                    If sumValue > 0 Then
                        documentsWithSum = documentsWithSum + 1
                        totalSum = totalSum + sumValue
                        Debug.Print "  Added to total: " & sumValue
                    End If
                End If
            End If
        End If
    Next i
    
    If documentsWithSum > 1 Then
        CalculateTotalDocumentsSum = 0
        Debug.Print "Multiple documents with sum - returning 0"
    ElseIf documentsWithSum = 1 Then
        CalculateTotalDocumentsSum = totalSum
        Debug.Print "Single document with sum: " & totalSum
    Else
        CalculateTotalDocumentsSum = 0
        Debug.Print "No documents with sum - returning 0"
    End If
    
    Debug.Print "=== DEBUG CalculateTotalDocumentsSum END, result: " & CalculateTotalDocumentsSum
End Function


' ======================================================================
'                    REMAINING FUNCTIONS (abbreviated)
' ======================================================================

Public Function FormatRecipientAddress(addressParts As Variant) As String
    Dim fullAddress As String
    Dim addressComponents As Collection
    Set addressComponents = New Collection
    
    Dim i As Integer
    For i = AddressIndexStreet To AddressIndexPhone
        If Len(Trim(CStr(addressParts(i)))) > 0 Then
            addressComponents.Add Trim(CStr(addressParts(i)))
        End If
    Next i
    
    For i = 1 To addressComponents.count
        If i > 1 Then fullAddress = fullAddress & ", "
        fullAddress = fullAddress & addressComponents(i)
    Next i
    
    FormatRecipientAddress = fullAddress
End Function

Public Function TryParseDate(rawText As String, ByRef outDate As Date) As Boolean
    Dim t As String, ok As Boolean
    Dim clean As String, i As Long, ch As String
    
    On Error Resume Next
    If TryParseDateExtended(rawText, outDate) Then
        TryParseDate = True
        Exit Function
    End If
    On Error GoTo 0
    
    TryParseDate = False
    
    If Len(Trim(rawText)) = 0 Then Exit Function
    
    On Error Resume Next
    If IsDate(rawText) Then
        outDate = CDate(rawText)
        TryParseDate = True
        Exit Function
    End If
    On Error GoTo 0
    
    t = Replace(rawText, "/", ".")
    
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If IsNumeric(ch) Then clean = clean & ch
    Next i
    
    Select Case Len(clean)
        Case 8
            ok = IsDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & "." & Right(clean, 4))
            If ok Then outDate = CDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & "." & Right(clean, 4))
        Case 6
            ok = IsDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & ".20" & Right(clean, 2))
            If ok Then outDate = CDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & ".20" & Right(clean, 2))
        Case 5
            ok = IsDate(Left(clean, 1) & "." & Mid(clean, 2, 2) & ".20" & Right(clean, 2))
            If ok Then outDate = CDate(Left(clean, 1) & "." & Mid(clean, 2, 2) & ".20" & Right(clean, 2))
        Case 4
            ok = IsDate(Left(clean, 2) & "." & Right(clean, 2) & "." & Year(Date))
            If ok Then outDate = CDate(Left(clean, 2) & "." & Right(clean, 2) & "." & Year(Date))
        Case Else
            ok = False
    End Select
    
    TryParseDate = ok
End Function

Public Function ResolveLetterDateOrToday(rawText As String) As Date
    If Len(Trim(rawText)) = 0 Then
        ResolveLetterDateOrToday = Date
        Exit Function
    End If
    
    If IsDate(rawText) Then
        ResolveLetterDateOrToday = CDate(rawText)
        Exit Function
    End If
    
    Dim parsedDate As Date
    If TryParseDate(rawText, parsedDate) Then
        ResolveLetterDateOrToday = parsedDate
    Else
        ResolveLetterDateOrToday = Date
    End If
End Function

Public Function HasAddressDataChanged(rowNumber As Long, newAddressArray As Variant) As Boolean
    HasAddressDataChanged = False
    
    On Error GoTo CompareError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    Dim i As Long
    For i = AddressIndexAddressee To AddressIndexPhone
        Dim sheetValue As String
        Dim formValue As String
        Dim columnNumber As AddressColumns
        
        columnNumber = AddressColumnFromIndex(i)
        sheetValue = UCase(Trim(CStr(ws.Cells(rowNumber, columnNumber).value)))
        formValue = UCase(Trim(CStr(newAddressArray(i))))
        
        If sheetValue <> formValue Then
            Debug.Print "Change in column " & columnNumber & ": '" & ws.Cells(rowNumber, columnNumber).value & "' -> '" & newAddressArray(i) & "'"
            HasAddressDataChanged = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
CompareError:
    Debug.Print "Error comparing address data: " & Err.description
    HasAddressDataChanged = False
End Function

Public Function FormatLetterDate(dateValue As String) As String
    On Error GoTo FormatError
    
    Dim d As Date
    
    If IsDate(dateValue) Then
        d = CDate(dateValue)
    Else
        If TryParseDateExtended(dateValue, d) Then
        Else
            FormatLetterDate = dateValue
            Exit Function
        End If
    End If
    
    Dim dayNum As Integer, monthNum As Integer, yearNum As Integer
    dayNum = Day(d)
    monthNum = Month(d)
    yearNum = Year(d)
    
    Dim monthName As String
    monthName = GetDirectMonthName(monthNum)
    
    FormatLetterDate = dayNum & " " & monthName & " " & yearNum
    
    Exit Function
    
FormatError:
    FormatLetterDate = dateValue
End Function

Private Function GetDirectMonthName(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetDirectMonthName = BuildUnicodeString(&H44F, &H43D, &H432, &H430, &H440, &H44F)
        Case 2: GetDirectMonthName = BuildUnicodeString(&H444, &H435, &H432, &H440, &H430, &H43B, &H44F)
        Case 3: GetDirectMonthName = BuildUnicodeString(&H43C, &H430, &H440, &H442, &H430)
        Case 4: GetDirectMonthName = BuildUnicodeString(&H430, &H43F, &H440, &H435, &H43B, &H44F)
        Case 5: GetDirectMonthName = BuildUnicodeString(&H43C, &H430, &H44F)
        Case 6: GetDirectMonthName = BuildUnicodeString(&H438, &H44E, &H43D, &H44F)
        Case 7: GetDirectMonthName = BuildUnicodeString(&H438, &H44E, &H43B, &H44F)
        Case 8: GetDirectMonthName = BuildUnicodeString(&H430, &H432, &H433, &H443, &H441, &H442, &H430)
        Case 9: GetDirectMonthName = BuildUnicodeString(&H441, &H435, &H43D, &H442, &H44F, &H431, &H440, &H44F)
        Case 10: GetDirectMonthName = BuildUnicodeString(&H43E, &H43A, &H442, &H44F, &H431, &H440, &H44F)
        Case 11: GetDirectMonthName = BuildUnicodeString(&H43D, &H43E, &H44F, &H431, &H440, &H44F)
        Case 12: GetDirectMonthName = BuildUnicodeString(&H434, &H435, &H43A, &H430, &H431, &H440, &H44F)
        Case Else: GetDirectMonthName = t("core.date.unknown_month", "unknown_month")
    End Select
End Function

Private Function BuildUnicodeString(ParamArray codePoints() As Variant) As String
    Dim i As Long

    BuildUnicodeString = ""
    For i = LBound(codePoints) To UBound(codePoints)
        BuildUnicodeString = BuildUnicodeString & ChrW(CLng(codePoints(i)))
    Next i
End Function

Public Sub ShowLetterCreatorDelayed()
    ShowLetterCreatorDeferred
End Sub

Public Sub ShowLetterCreatorDeferred()
    On Error GoTo DelayedErrorHandler
    
    Load frmLetterCreator
    frmLetterCreator.Show vbModeless
    Exit Sub
    
DelayedErrorHandler:
    MsgBox t("core.form.open_creator_error", "Failed to open letter creation form: ") & Err.description, vbCritical
End Sub

Public Sub StartFormirovanieLetters()
    ShowLetterCreator
End Sub

Public Sub ShowLetterCreator()
    Load frmLetterCreator
    frmLetterCreator.Show vbModeless
End Sub

Public Function GetDocumentTypeText(documentType As String) As String
    If NormalizeDocumentTypeKey(documentType) = DocumentTypeKeyOwnConfirmation Then
        GetDocumentTypeText = t("core.letter.text.own_confirmation", "forwarding the following documents to your address for confirmation")
    Else
        GetDocumentTypeText = t("core.letter.text.confirmed", "forwarding confirmed accounting documents to your address")
    End If
    
    On Error GoTo ReadTextError
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    
    If ws Is Nothing Then Exit Function
    
    Dim tbl As ListObject
    Set tbl = GetLetterTextsTable(ws)
    
    If Not tbl Is Nothing Then
        If tbl.ListRows.count >= 1 Then
            Dim textResult As String
            If NormalizeDocumentTypeKey(documentType) = DocumentTypeKeyOwnConfirmation Then
                textResult = Trim(tbl.DataBodyRange.Cells(TextTableRowOwnForConfirmation, TextTableColumnBody).value)
            ElseIf tbl.ListRows.count >= TextTableRowConfirmedDocuments Then
                textResult = Trim(tbl.DataBodyRange.Cells(TextTableRowConfirmedDocuments, TextTableColumnBody).value)
            End If
            
            If Len(textResult) > 0 Then
                textResult = LCase(Left(textResult, 1)) & Mid(textResult, 2)
                GetDocumentTypeText = textResult
            End If
        End If
    End If
    
    Exit Function
    
ReadTextError:
    Debug.Print "GetDocumentTypeText fallback used: " & Err.description
End Function

Public Function BuildHistoryLoadedCaption(letterCount As Long) As String
    BuildHistoryLoadedCaption = t("form.letter_history.msg.letters_loaded", "Letters loaded: ") & letterCount
End Function

Public Function BuildHistoryShowingAllCaption(letterCount As Long) As String
    BuildHistoryShowingAllCaption = t("form.letter_history.msg.showing_all", "Showing all letters: ") & letterCount
End Function

Public Function BuildHistoryAmountSearchCaption(searchText As String) As String
    BuildHistoryAmountSearchCaption = t("form.letter_history.msg.searching_amount", "Searching for number ") & searchText & t("form.letter_history.msg.searching_amount_suffix", " in document amounts...")
End Function

Public Function BuildHistoryFoundCaption(foundCount As Long, totalCount As Long) As String
    BuildHistoryFoundCaption = t("form.letter_history.msg.letters_found", "Letters found: ") & foundCount & t("form.letter_history.msg.out_of", " of ") & totalCount
End Function

Private Function GetLetterTextsTable(ws As Worksheet) As ListObject
    On Error Resume Next
    Set GetLetterTextsTable = ws.ListObjects(LetterTextsTableName)
    If GetLetterTextsTable Is Nothing Then
        Set GetLetterTextsTable = ws.ListObjects(LegacyLetterTextsTableName)
    End If
    If GetLetterTextsTable Is Nothing Then
        Set GetLetterTextsTable = ws.ListObjects(LegacyLetterTextsTableNameLocalized)
    End If
    On Error GoTo 0
End Function



Public Sub SafeReplaceInWord(wordDoc As Object, findText As String, replaceText As String)
    WordInteropSafeReplaceInWord wordDoc, findText, replaceText
End Sub

Public Sub SafeReplaceInWordWithFragments(wordDoc As Object, findText As String, fragments As Collection)
    WordInteropSafeReplaceInWordWithFragments wordDoc, findText, fragments
End Sub

Public Sub FormatAttachmentsInWord(rng As Object, Optional fontSize As Integer = 10)
    WordInteropFormatAttachmentsInWord rng, fontSize
End Sub

Public Function GenerateFileNameWithExecutor(addressee As String, letterNumber As String, executor As String) As String
    Dim cleanAddressee As String, cleanNumber As String, cleanExecutor As String
    Dim currentDate As String
    
    cleanAddressee = CleanFileName(addressee)
    cleanNumber = CleanFileName(letterNumber)
    cleanExecutor = CleanFileName(executor)
    currentDate = Format(Date, "dd.mm.yyyy")
    
    GenerateFileNameWithExecutor = ThisWorkbook.Path & "\" & cleanAddressee & "_" & _
                                  cleanNumber & "_" & currentDate & "_" & cleanExecutor & ".docx"
End Function

Public Function CleanFileName(inputName As String) As String
    Dim result As String
    result = Trim(inputName)
    
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    result = Replace(result, " ", "_")
    
    If Len(result) > 30 Then result = Left(result, 30)
    
    CleanFileName = result
End Function

Public Sub ClearHighlight()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not ws Is Nothing Then
        ws.Cells.Interior.Pattern = xlNone
        ws.Cells.Interior.ColorIndex = xlNone
        Debug.Print "Row highlight cleared"
        
        Application.StatusBar = False
    End If
    
    On Error GoTo 0
End Sub

Public Sub RestoreFocusToHistory()
    Dim historyForm As Object
    If Not TryGetLoadedUserForm("frmLetterHistory", historyForm) Then Exit Sub
    
    If Not historyForm Is Nothing Then
        historyForm.SetFocus
        historyForm.ZOrder 0
        Debug.Print "Focus returned to letter history form"
    End If
End Sub

Public Sub ShowLetterHistoryModeless()
    On Error GoTo ShowHistoryError
    
    Dim existingForm As Object
    Call TryGetLoadedUserForm("frmLetterHistory", existingForm)
    
    If Not existingForm Is Nothing Then
        existingForm.SetFocus
        existingForm.ZOrder 0
        MsgBox t("form.letter_history.msg.already_open", "Letter history form is already open!"), vbInformation
    Else
        Load frmLetterHistory
        frmLetterHistory.Show vbModeless
        Debug.Print "Letter history form launched modelessly from ModuleMain"
    End If
    
    Exit Sub
    
ShowHistoryError:
    MsgBox t("form.letter_history.msg.open_error", "Error opening letter history form: ") & Err.description, vbCritical
End Sub

Public Sub ClearAddressHighlight()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    If Not ws Is Nothing Then
        ws.Cells.Interior.Pattern = xlNone
        Debug.Print "Address row highlight cleared"
    End If
    
    On Error GoTo 0
End Sub

Public Sub ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = False
    Debug.Print "Excel status bar cleared"
    On Error GoTo 0
End Sub

Public Sub SetStatusBarMessage(message As String, Optional clearAfterSeconds As Integer = 0)
    On Error Resume Next
    
    Application.StatusBar = message
    Debug.Print "Status bar: " & message
    
    If clearAfterSeconds > 0 Then
        Application.OnTime Now + TimeValue("00:00:" & Format(clearAfterSeconds, "00")), "ClearStatusBar"
    End If
    
    On Error GoTo 0
End Sub



