Attribute VB_Name = "ModuleDispatchRepository"
' ======================================================================
' Module: ModuleDispatchRepository
' Author: CreateLetter contributors
' Purpose: Workbook repository helpers for envelope formats, senders, and dispatch items
' Version: 1.0.0 - 26.04.2026
' ======================================================================

Option Explicit

Public Function DispatchRepositoryLoadEnvelopeFormats() As Collection
    Set DispatchRepositoryLoadEnvelopeFormats = New Collection

    On Error GoTo LoadError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("EnvelopeFormats")

    Dim formatData As Variant
    formatData = RepositoryReadWorksheetMatrix(ws, EnvelopeFormatColumnKey, EnvelopeFormatColumnSortOrder, EnvelopeFormatsTableName)
    If IsEmpty(formatData) Then Exit Function

    Dim i As Long
    For i = LBound(formatData, 1) To UBound(formatData, 1)
        If DispatchRepositoryIsTruthy(formatData(i, EnvelopeFormatColumnIsActive)) Then
            DispatchRepositoryLoadEnvelopeFormats.Add DispatchRepositoryCreateEnvelopeFormatDescriptor( _
                CStr(formatData(i, EnvelopeFormatColumnKey)), _
                CStr(formatData(i, EnvelopeFormatColumnDisplayName)), _
                DispatchRepositoryIsTruthy(formatData(i, EnvelopeFormatColumnIsActive)), _
                CLng(Val(CStr(formatData(i, EnvelopeFormatColumnSortOrder)))))
        End If
    Next i

    Exit Function

LoadError:
    Debug.Print "DispatchRepositoryLoadEnvelopeFormats error: " & Err.description
End Function

Public Function DispatchRepositoryLoadSenders() As Collection
    Set DispatchRepositoryLoadSenders = New Collection

    On Error GoTo LoadError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Senders")

    Dim senderData As Variant
    senderData = RepositoryReadWorksheetMatrix(ws, SenderColumnName, SenderColumnIsDefault, SendersTableName)
    If IsEmpty(senderData) Then Exit Function

    Dim i As Long
    For i = LBound(senderData, 1) To UBound(senderData, 1)
        If Len(Trim$(CStr(senderData(i, SenderColumnName)))) > 0 Then
            DispatchRepositoryLoadSenders.Add DispatchRepositoryCreateSenderDescriptor( _
                CStr(senderData(i, SenderColumnName)), _
                CStr(senderData(i, SenderColumnAddressLine1)), _
                CStr(senderData(i, SenderColumnAddressLine2)), _
                CStr(senderData(i, SenderColumnAddressLine3)), _
                CStr(senderData(i, SenderColumnPostalCode)), _
                CStr(senderData(i, SenderColumnPhone)), _
                DispatchRepositoryIsTruthy(senderData(i, SenderColumnIsDefault)))
        End If
    Next i

    Exit Function

LoadError:
    Debug.Print "DispatchRepositoryLoadSenders error: " & Err.description
End Function

Public Function DispatchRepositoryGetDefaultSenderName() As String
    On Error GoTo LoadError

    Dim senders As Collection
    Set senders = DispatchRepositoryLoadSenders()

    Dim i As Long
    For i = 1 To senders.count
        If DispatchRepositoryIsSenderDefault(senders(i)) Then
            DispatchRepositoryGetDefaultSenderName = DispatchRepositoryGetSenderName(senders(i))
            Exit Function
        End If
    Next i

    If senders.count > 0 Then
        DispatchRepositoryGetDefaultSenderName = DispatchRepositoryGetSenderName(senders(1))
    End If

    Exit Function

LoadError:
    Debug.Print "DispatchRepositoryGetDefaultSenderName error: " & Err.description
End Function

Public Function DispatchRepositoryGetEnvelopeFormatDisplay(envelopeFormatKey As String) As String
    Dim normalizedKey As String
    normalizedKey = LCase$(Trim$(envelopeFormatKey))

    If Len(normalizedKey) = 0 Then Exit Function

    Select Case normalizedKey
    Case "c4", "c5", "dl"
        DispatchRepositoryGetEnvelopeFormatDisplay = t("dispatch.envelope_format." & normalizedKey, UCase$(normalizedKey))
    Case Else
        DispatchRepositoryGetEnvelopeFormatDisplay = UCase$(normalizedKey)
    End Select
End Function

Public Function DispatchRepositoryCreateItemFromLetterFields( _
    ByVal addressee As String, _
    ByVal letterNumber As String, _
    ByVal letterDate As String, _
    ByVal senderName As String, _
    ByVal envelopeFormatKey As String, _
    Optional ByVal mailType As String = "", _
    Optional ByVal mass As String = "", _
    Optional ByVal declaredValue As String = "", _
    Optional ByVal comment As String = "", _
    Optional ByVal phone As String = "", _
    Optional ByVal batchId As String = "", _
    Optional ByVal status As String = "") As String

    On Error GoTo SaveError

    Dim addressLine As String
    Dim postalCode As String
    Dim resolvedPhone As String
    Call DispatchRepositoryTryResolveAddressByAddressee(addressee, addressLine, postalCode, resolvedPhone)

    If Len(Trim$(phone)) = 0 Then
        phone = resolvedPhone
    End If

    Dim dispatchId As String
    dispatchId = DispatchRepositoryGenerateId(letterNumber)

    Dim dispatchTable As ListObject
    Set dispatchTable = DispatchRepositoryGetTable("DispatchItems", DispatchItemsTableName)

    Dim newRow As ListRow
    Set newRow = dispatchTable.ListRows.Add

    With newRow.Range
        .Cells(1, DispatchItemColumnId).value = dispatchId
        .Cells(1, DispatchItemColumnLetterNumber).value = letterNumber
        .Cells(1, DispatchItemColumnLetterDate).value = letterDate
        .Cells(1, DispatchItemColumnAddressee).value = addressee
        .Cells(1, DispatchItemColumnAddressLine).value = addressLine
        .Cells(1, DispatchItemColumnPostalCode).value = postalCode
        .Cells(1, DispatchItemColumnSenderName).value = senderName
        .Cells(1, DispatchItemColumnEnvelopeFormatKey).value = LCase$(Trim$(envelopeFormatKey))
        .Cells(1, DispatchItemColumnMailType).value = mailType
        .Cells(1, DispatchItemColumnMass).value = mass
        .Cells(1, DispatchItemColumnDeclaredValue).value = declaredValue
        .Cells(1, DispatchItemColumnComment).value = comment
        .Cells(1, DispatchItemColumnPhone).value = phone
        .Cells(1, DispatchItemColumnBatchId).value = batchId
        .Cells(1, DispatchItemColumnStatus).value = DispatchRepositoryResolveDispatchStatus(status)
        .Cells(1, DispatchItemColumnCreatedAt).value = Format$(Now, "dd.mm.yyyy hh:nn:ss")
    End With

    DispatchRepositoryCreateItemFromLetterFields = dispatchId
    Exit Function

SaveError:
    Debug.Print "DispatchRepositoryCreateItemFromLetterFields error: " & Err.description
    DispatchRepositoryCreateItemFromLetterFields = ""
End Function

Public Function DispatchRepositoryCreateItemFromHistoryRecord( _
    record As clsLetterHistoryRecord, _
    ByVal senderName As String, _
    ByVal envelopeFormatKey As String, _
    Optional ByVal mailType As String = "", _
    Optional ByVal mass As String = "", _
    Optional ByVal declaredValue As String = "", _
    Optional ByVal comment As String = "", _
    Optional ByVal phone As String = "", _
    Optional ByVal batchId As String = "", _
    Optional ByVal status As String = "") As String

    If record Is Nothing Then Exit Function

    DispatchRepositoryCreateItemFromHistoryRecord = DispatchRepositoryCreateItemFromLetterFields( _
        record.Addressee, _
        record.OutgoingNumber, _
        record.OutgoingDate, _
        senderName, _
        envelopeFormatKey, _
        mailType, _
        mass, _
        declaredValue, _
        comment, _
        phone, _
        batchId, _
        status)
End Function

Public Function DispatchRepositoryLoadDispatchItems() As Collection
    Set DispatchRepositoryLoadDispatchItems = New Collection

    On Error GoTo LoadError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DispatchItems")

    Dim dispatchData As Variant
    dispatchData = RepositoryReadWorksheetMatrix(ws, DispatchItemColumnId, DispatchItemColumnCreatedAt, DispatchItemsTableName)
    If IsEmpty(dispatchData) Then Exit Function

    Dim i As Long
    For i = LBound(dispatchData, 1) To UBound(dispatchData, 1)
        If Len(Trim$(CStr(dispatchData(i, DispatchItemColumnId)))) > 0 Then
            DispatchRepositoryLoadDispatchItems.Add DispatchRepositoryCreateDispatchItemDescriptor(dispatchData, i)
        End If
    Next i

    Exit Function

LoadError:
    Debug.Print "DispatchRepositoryLoadDispatchItems error: " & Err.description
End Function

Private Function DispatchRepositoryGetTable(sheetName As String, tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set DispatchRepositoryGetTable = ws.ListObjects.Item(tableName)
End Function

Private Function DispatchRepositoryCreateEnvelopeFormatDescriptor( _
    ByVal formatKey As String, _
    ByVal displayName As String, _
    ByVal isActive As Boolean, _
    ByVal sortOrder As Long) As Variant

    Dim descriptor(1 To 4) As Variant
    descriptor(EnvelopeFormatColumnKey) = LCase$(Trim$(formatKey))
    descriptor(EnvelopeFormatColumnDisplayName) = displayName
    descriptor(EnvelopeFormatColumnIsActive) = isActive
    descriptor(EnvelopeFormatColumnSortOrder) = sortOrder
    DispatchRepositoryCreateEnvelopeFormatDescriptor = descriptor
End Function

Private Function DispatchRepositoryCreateSenderDescriptor( _
    ByVal senderName As String, _
    ByVal addressLine1 As String, _
    ByVal addressLine2 As String, _
    ByVal addressLine3 As String, _
    ByVal postalCode As String, _
    ByVal phone As String, _
    ByVal isDefault As Boolean) As Variant

    Dim descriptor(1 To 7) As Variant
    descriptor(SenderColumnName) = senderName
    descriptor(SenderColumnAddressLine1) = addressLine1
    descriptor(SenderColumnAddressLine2) = addressLine2
    descriptor(SenderColumnAddressLine3) = addressLine3
    descriptor(SenderColumnPostalCode) = postalCode
    descriptor(SenderColumnPhone) = phone
    descriptor(SenderColumnIsDefault) = isDefault
    DispatchRepositoryCreateSenderDescriptor = descriptor
End Function

Private Function DispatchRepositoryCreateDispatchItemDescriptor(dispatchData As Variant, rowIndex As Long) As Variant
    Dim descriptor(1 To DispatchItemColumnCreatedAt) As Variant
    Dim columnIndex As Long

    For columnIndex = DispatchItemColumnId To DispatchItemColumnCreatedAt
        descriptor(columnIndex) = dispatchData(rowIndex, columnIndex)
    Next columnIndex

    DispatchRepositoryCreateDispatchItemDescriptor = descriptor
End Function

Private Function DispatchRepositoryGenerateId(letterNumber As String) As String
    Dim normalizedNumber As String
    normalizedNumber = Replace(Replace(Trim$(letterNumber), "/", "-"), " ", "")

    If Len(normalizedNumber) = 0 Then
        normalizedNumber = "dispatch"
    End If

    DispatchRepositoryGenerateId = "dispatch-" & Format$(Now, "yyyymmddhhnnss") & "-" & normalizedNumber
End Function

Private Function DispatchRepositoryResolveDispatchStatus(status As String) As String
    If Len(Trim$(status)) = 0 Then
        DispatchRepositoryResolveDispatchStatus = "draft"
    Else
        DispatchRepositoryResolveDispatchStatus = LCase$(Trim$(status))
    End If
End Function

Private Function DispatchRepositoryTryResolveAddressByAddressee( _
    ByVal addressee As String, _
    ByRef addressLine As String, _
    ByRef postalCode As String, _
    ByRef phone As String) As Boolean

    DispatchRepositoryTryResolveAddressByAddressee = False

    On Error GoTo ResolveError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")

    Dim addressData As Variant
    addressData = RepositoryReadWorksheetMatrix(ws, AddressColumnAddressee, AddressColumnGroup, AddressesTableName)
    If IsEmpty(addressData) Then Exit Function

    Dim normalizedAddressee As String
    normalizedAddressee = UCase$(Trim$(addressee))

    Dim i As Long
    For i = LBound(addressData, 1) To UBound(addressData, 1)
        If UCase$(Trim$(CStr(addressData(i, AddressColumnAddressee)))) = normalizedAddressee Then
            Dim addressParts(AddressIndexGroup) As String
            addressParts(AddressIndexAddressee) = CStr(addressData(i, AddressColumnAddressee))
            addressParts(AddressIndexStreet) = CStr(addressData(i, AddressColumnStreet))
            addressParts(AddressIndexCity) = CStr(addressData(i, AddressColumnCity))
            addressParts(AddressIndexDistrict) = CStr(addressData(i, AddressColumnDistrict))
            addressParts(AddressIndexRegion) = CStr(addressData(i, AddressColumnRegion))
            addressParts(AddressIndexPostalCode) = CStr(addressData(i, AddressColumnPostalCode))
            addressParts(AddressIndexPhone) = CStr(addressData(i, AddressColumnPhone))
            addressParts(AddressIndexGroup) = CStr(addressData(i, AddressColumnGroup))

            addressLine = FormatRecipientAddress(addressParts)
            postalCode = addressParts(AddressIndexPostalCode)
            phone = addressParts(AddressIndexPhone)
            DispatchRepositoryTryResolveAddressByAddressee = True
            Exit Function
        End If
    Next i

    Exit Function

ResolveError:
    DispatchRepositoryTryResolveAddressByAddressee = False
End Function

Private Function DispatchRepositoryIsTruthy(value As Variant) As Boolean
    Select Case VarType(value)
    Case vbBoolean
        DispatchRepositoryIsTruthy = CBool(value)
    Case vbString
        Dim normalizedText As String
        normalizedText = UCase$(Trim$(CStr(value)))
        DispatchRepositoryIsTruthy = (normalizedText = "TRUE" Or normalizedText = "YES" Or normalizedText = "1" Or normalizedText = "ДА")
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        DispatchRepositoryIsTruthy = (Val(CStr(value)) <> 0)
    Case Else
        DispatchRepositoryIsTruthy = False
    End Select
End Function

Private Function DispatchRepositoryIsSenderDefault(senderDescriptor As Variant) As Boolean
    On Error GoTo NotDefault
    DispatchRepositoryIsSenderDefault = CBool(senderDescriptor(SenderColumnIsDefault))
    Exit Function
NotDefault:
    DispatchRepositoryIsSenderDefault = False
End Function

Private Function DispatchRepositoryGetSenderName(senderDescriptor As Variant) As String
    On Error GoTo MissingName
    DispatchRepositoryGetSenderName = CStr(senderDescriptor(SenderColumnName))
    Exit Function
MissingName:
    DispatchRepositoryGetSenderName = ""
End Function
