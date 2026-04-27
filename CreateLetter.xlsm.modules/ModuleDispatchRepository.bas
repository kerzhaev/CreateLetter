Attribute VB_Name = "ModuleDispatchRepository"

' ======================================================================

' Module: ModuleDispatchRepository

' Author: CreateLetter contributors

' Purpose: Workbook repository helpers for envelope formats, senders, dispatch packages, and registry metadata

' Version: 1.1.0 - 26.04.2026

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

            Dim formatDescriptor As Variant

            formatDescriptor = DispatchRepositoryCreateEnvelopeFormatDescriptor(CStr(formatData(i, EnvelopeFormatColumnKey)), CStr(formatData(i, EnvelopeFormatColumnDisplayName)), DispatchRepositoryIsTruthy(formatData(i, EnvelopeFormatColumnIsActive)), CLng(Val(CStr(formatData(i, EnvelopeFormatColumnSortOrder)))))

            DispatchRepositoryLoadEnvelopeFormats.Add formatDescriptor

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

            Dim senderDescriptor As Variant

            senderDescriptor = DispatchRepositoryCreateSenderDescriptor(CStr(senderData(i, SenderColumnName)), CStr(senderData(i, SenderColumnAddressLine1)), CStr(senderData(i, SenderColumnAddressLine2)), CStr(senderData(i, SenderColumnAddressLine3)), CStr(senderData(i, SenderColumnPostalCode)), CStr(senderData(i, SenderColumnPhone)), DispatchRepositoryIsTruthy(senderData(i, SenderColumnIsDefault)))

            DispatchRepositoryLoadSenders.Add senderDescriptor

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



Public Function DispatchRepositoryGetSenderPostalCode(senderName As String) As String

    On Error GoTo LookupError



    Dim senderDescriptor As Variant

    If DispatchRepositoryTryGetSenderDescriptor(senderName, senderDescriptor) Then

        DispatchRepositoryGetSenderPostalCode = CStr(senderDescriptor(SenderColumnPostalCode))

    End If



    Exit Function



LookupError:

    DispatchRepositoryGetSenderPostalCode = ""

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



Public Function DispatchRepositoryBuildRecipientPreviewByAddressee(ByVal Addressee As String) As String

    On Error GoTo BuildError



    Dim addressLine As String

    Dim postalCode As String

    Dim phone As String



    If DispatchRepositoryTryResolveAddressByAddressee(Addressee, addressLine, postalCode, phone) Then

        DispatchRepositoryBuildRecipientPreviewByAddressee = Addressee

        If Len(Trim$(addressLine)) > 0 Then

            DispatchRepositoryBuildRecipientPreviewByAddressee = DispatchRepositoryBuildRecipientPreviewByAddressee & vbCrLf & addressLine

        End If

        Exit Function

    End If



    DispatchRepositoryBuildRecipientPreviewByAddressee = Addressee

    Exit Function



BuildError:

    DispatchRepositoryBuildRecipientPreviewByAddressee = Addressee

End Function



Public Function DispatchRepositoryGetQueuedLetterKeySet() As Object

    Dim queuedLetterKeys As Object

    Set queuedLetterKeys = CreateObject("Scripting.Dictionary")

    queuedLetterKeys.CompareMode = vbTextCompare



    On Error GoTo LoadError



    Dim dispatchItems As Collection

    Set dispatchItems = DispatchRepositoryLoadDispatchItems()

    If dispatchItems Is Nothing Then

        Set DispatchRepositoryGetQueuedLetterKeySet = queuedLetterKeys

        Exit Function

    End If



    Dim i As Long

    For i = 1 To dispatchItems.count

        Dim itemDescriptor As Variant

        itemDescriptor = dispatchItems(i)



        If Len(Trim$(CStr(itemDescriptor(DispatchItemColumnLetterNumber)))) > 0 Then

            queuedLetterKeys(DispatchRepositoryBuildHistoryKey( _

                CStr(itemDescriptor(DispatchItemColumnAddressee)), _

                CStr(itemDescriptor(DispatchItemColumnLetterNumber)), _

                CStr(itemDescriptor(DispatchItemColumnLetterDate)))) = True

        End If

    Next i



    Set DispatchRepositoryGetQueuedLetterKeySet = queuedLetterKeys

    Exit Function



LoadError:

    Set DispatchRepositoryGetQueuedLetterKeySet = CreateObject("Scripting.Dictionary")

    DispatchRepositoryGetQueuedLetterKeySet.CompareMode = vbTextCompare

End Function



Public Function DispatchRepositoryCreateItemFromLetterFields( _

    ByVal Addressee As String, _

    ByVal letterNumber As String, _

    ByVal letterDate As String, _

    ByVal letterRowNumber As Long, _

    ByVal senderName As String, _

    ByVal envelopeFormatKey As String, _

    Optional ByVal mailType As String = "", _

    Optional ByVal mass As String = "", _

    Optional ByVal declaredValue As String = "", _

    Optional ByVal comment As String = "", _

    Optional ByVal phone As String = "", _

    Optional ByVal batchId As String = "", _

    Optional ByVal status As String = "", _

    Optional ByVal registryNumber As String = "", _

    Optional ByVal registryDate As String = "") As String



    On Error GoTo SaveError



    Dim addressLine As String

    Dim postalCode As String

    Dim resolvedPhone As String

    Call DispatchRepositoryTryResolveAddressByAddressee(Addressee, addressLine, postalCode, resolvedPhone)



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

        .Cells(1, DispatchItemColumnLetterRowNumber).value = letterRowNumber

        .Cells(1, DispatchItemColumnAddressee).value = Addressee

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

        .Cells(1, DispatchItemColumnRegistryNumber).value = registryNumber

        .Cells(1, DispatchItemColumnRegistryDate).value = registryDate

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

    Optional ByVal status As String = "", _

    Optional ByVal registryNumber As String = "", _

    Optional ByVal registryDate As String = "") As String



    If record Is Nothing Then Exit Function



    DispatchRepositoryCreateItemFromHistoryRecord = DispatchRepositoryCreateItemFromLetterFields( _

        record.Addressee, _

        record.OutgoingNumber, _

        record.OutgoingDate, _

        record.rowNumber, _

        senderName, _

        envelopeFormatKey, _

        mailType, _

        mass, _

        declaredValue, _

        comment, _

        phone, _

        batchId, _

        status, _

        registryNumber, _

        registryDate)

End Function



Public Function DispatchRepositoryCreatePackageFromHistoryRecords( _

    packageRecords As Collection, _

    ByVal senderName As String, _

    ByVal envelopeFormatKey As String, _

    ByVal registryNumber As String, _

    ByVal registryDate As String, _

    Optional ByVal mailType As String = "", _

    Optional ByVal mass As String = "", _

    Optional ByVal declaredValue As String = "", _

    Optional ByVal comment As String = "", _

    Optional ByVal phone As String = "") As String



    On Error GoTo SaveError



    If packageRecords Is Nothing Then Exit Function

    If packageRecords.count = 0 Then Exit Function



    Dim firstRecord As clsLetterHistoryRecord

    Set firstRecord = packageRecords(1)

    If firstRecord Is Nothing Then Exit Function



    Dim batchId As String

    batchId = DispatchRepositoryGenerateBatchId(firstRecord.Addressee)



    Dim i As Long

    For i = 1 To packageRecords.count

        Dim record As clsLetterHistoryRecord

        Set record = packageRecords(i)



        If record Is Nothing Then GoTo NextRecord



        If Len(DispatchRepositoryCreateItemFromHistoryRecord( _

            record, _

            senderName, _

            envelopeFormatKey, _

            mailType, _

            mass, _

            declaredValue, _

            comment, _

            phone, _

            batchId, _

            "queued", _

            registryNumber, _

            registryDate)) = 0 Then

            Err.Raise vbObjectError + 4600, "DispatchRepositoryCreatePackageFromHistoryRecords", "Failed to save a dispatch item into the package."

        End If



NextRecord:

    Next i



    DispatchRepositoryCreatePackageFromHistoryRecords = batchId

    Exit Function



SaveError:

    Debug.Print "DispatchRepositoryCreatePackageFromHistoryRecords error: " & Err.description

    DispatchRepositoryCreatePackageFromHistoryRecords = ""

End Function



Public Function DispatchRepositoryLoadDispatchItems() As Collection

    Set DispatchRepositoryLoadDispatchItems = New Collection



    On Error GoTo LoadError



    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("DispatchItems")



    Dim dispatchData As Variant

    dispatchData = RepositoryReadWorksheetMatrix(ws, DispatchItemColumnId, DispatchItemColumnRegistryDate, DispatchItemsTableName)

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



Public Sub DispatchRepositoryUpdateBatchStatus(ByVal batchId As String, ByVal status As String)

    On Error GoTo UpdateError



    If Len(Trim$(batchId)) = 0 Then Exit Sub



    Dim dispatchTable As ListObject

    Set dispatchTable = DispatchRepositoryGetTable("DispatchItems", DispatchItemsTableName)

    If dispatchTable.DataBodyRange Is Nothing Then Exit Sub



    Dim normalizedStatus As String

    normalizedStatus = DispatchRepositoryResolveDispatchStatus(status)



    Dim rowIndex As Long

    For rowIndex = 1 To dispatchTable.DataBodyRange.Rows.count

        If StrComp(CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnBatchId).value), batchId, vbTextCompare) = 0 Then

            dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnStatus).value = normalizedStatus

        End If

    Next rowIndex



    Exit Sub



UpdateError:

    Debug.Print "DispatchRepositoryUpdateBatchStatus error: " & Err.description

End Sub



Private Function DispatchRepositoryGetTable(sheetName As String, tableName As String) As ListObject

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(sheetName)

    Set DispatchRepositoryGetTable = ws.ListObjects.item(tableName)

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

    Dim descriptor(1 To DispatchItemColumnRegistryDate) As Variant

    Dim columnIndex As Long



    For columnIndex = DispatchItemColumnId To DispatchItemColumnRegistryDate

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



Private Function DispatchRepositoryGenerateBatchId(Addressee As String) As String

    Dim normalizedAddressee As String

    normalizedAddressee = Replace(Replace(Trim$(Addressee), " ", "-"), "/", "-")



    If Len(normalizedAddressee) = 0 Then

        normalizedAddressee = "batch"

    End If



    DispatchRepositoryGenerateBatchId = "batch-" & Format$(Now, "yyyymmddhhnnss") & "-" & Left$(normalizedAddressee, 30)

End Function



Private Function DispatchRepositoryResolveDispatchStatus(status As String) As String

    If Len(Trim$(status)) = 0 Then

        DispatchRepositoryResolveDispatchStatus = "draft"

    Else

        DispatchRepositoryResolveDispatchStatus = LCase$(Trim$(status))

    End If

End Function



Private Function DispatchRepositoryBuildHistoryKey( _

    ByVal Addressee As String, _

    ByVal letterNumber As String, _

    ByVal letterDate As String) As String



    DispatchRepositoryBuildHistoryKey = UCase$(Trim$(Addressee)) & "|" & UCase$(Trim$(letterNumber)) & "|" & UCase$(Trim$(letterDate))

End Function



Private Function DispatchRepositoryTryResolveAddressByAddressee( _

    ByVal Addressee As String, _

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

    normalizedAddressee = UCase$(Trim$(Addressee))



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



Private Function DispatchRepositoryTryGetSenderDescriptor(senderName As String, ByRef senderDescriptor As Variant) As Boolean

    DispatchRepositoryTryGetSenderDescriptor = False



    On Error GoTo LookupError



    Dim senders As Collection

    Set senders = DispatchRepositoryLoadSenders()



    Dim normalizedSenderName As String

    normalizedSenderName = UCase$(Trim$(senderName))



    Dim i As Long

    For i = 1 To senders.count

        If UCase$(Trim$(CStr(senders(i)(SenderColumnName)))) = normalizedSenderName Then

            senderDescriptor = senders(i)

            DispatchRepositoryTryGetSenderDescriptor = True

            Exit Function

        End If

    Next i



    Exit Function



LookupError:

    DispatchRepositoryTryGetSenderDescriptor = False

End Function



