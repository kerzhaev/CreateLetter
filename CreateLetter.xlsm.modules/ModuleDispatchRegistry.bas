Attribute VB_Name = "ModuleDispatchRegistry"
' ======================================================================
' Module: ModuleDispatchRegistry
' Author: CreateLetter contributors
' Purpose: Build and refresh the internal Excel dispatch registry from grouped dispatch packages
' Version: 1.2.0 - 27.04.2026
' ======================================================================

Option Explicit

Public Function BuildDispatchRegistryFromDispatchItems() As Long
    On Error GoTo BuildError

    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()

    ClearDispatchRegistry

    If dispatchItems Is Nothing Or dispatchItems.count = 0 Then Exit Function

    Dim groupedBatches As Object
    Set groupedBatches = GroupDispatchItemsByBatch(dispatchItems)
    If groupedBatches Is Nothing Then Exit Function
    If groupedBatches.count = 0 Then Exit Function

    Dim registryTable As ListObject
    Set registryTable = GetDispatchRegistryTable()

    Dim batchKey As Variant
    For Each batchKey In groupedBatches.keys
        Dim batchItems As Collection
        Set batchItems = groupedBatches(batchKey)
        If Not batchItems Is Nothing Then
            If batchItems.count > 0 Then
                AppendDispatchRegistryRow registryTable, batchItems
                UpdateLettersDispatchTracking batchItems
                DispatchRepositoryUpdateBatchStatus CStr(batchKey), "registered"
                BuildDispatchRegistryFromDispatchItems = BuildDispatchRegistryFromDispatchItems + 1
            End If
        End If
    Next batchKey

    Exit Function

BuildError:
    Debug.Print "BuildDispatchRegistryFromDispatchItems error: " & Err.description
    BuildDispatchRegistryFromDispatchItems = 0
End Function

Public Sub ClearDispatchRegistry()
    On Error GoTo ClearError

    Dim registryTable As ListObject
    Set registryTable = GetDispatchRegistryTable()

    If Not registryTable.DataBodyRange Is Nothing Then
        registryTable.DataBodyRange.Delete
    End If

    Exit Sub

ClearError:
    Debug.Print "ClearDispatchRegistry error: " & Err.description
End Sub

Private Function GetDispatchRegistryTable() As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DispatchRegistry")
    Set GetDispatchRegistryTable = ws.ListObjects.item(DispatchRegistryTableName)
End Function

Private Function GroupDispatchItemsByBatch(dispatchItems As Collection) As Object
    Dim groupedBatches As Object
    Set groupedBatches = CreateObject("Scripting.Dictionary")
    groupedBatches.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To dispatchItems.count
        Dim dispatchItem As Variant
        dispatchItem = dispatchItems(i)

        Dim batchKey As String
        batchKey = Trim$(CStr(dispatchItem(DispatchItemColumnBatchId)))
        If Len(batchKey) = 0 Then
            batchKey = CStr(dispatchItem(DispatchItemColumnId))
        End If

        If Not groupedBatches.Exists(batchKey) Then
            groupedBatches.Add batchKey, New Collection
        End If

        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)
        batchItems.Add dispatchItem
    Next i

    Set GroupDispatchItemsByBatch = groupedBatches
End Function

Private Sub AppendDispatchRegistryRow(registryTable As ListObject, batchItems As Collection)
    Dim newRow As ListRow
    Set newRow = registryTable.ListRows.Add

    Dim firstItem As Variant
    firstItem = batchItems(1)

    With newRow.Range
        .Cells(1, DispatchRegistryColumnRegistryNumber).value = CStr(firstItem(DispatchItemColumnRegistryNumber))
        .Cells(1, DispatchRegistryColumnRegistryDate).value = CStr(firstItem(DispatchItemColumnRegistryDate))
        .Cells(1, DispatchRegistryColumnBatchId).value = CStr(firstItem(DispatchItemColumnBatchId))
        .Cells(1, DispatchRegistryColumnAddressee).value = CStr(firstItem(DispatchItemColumnAddressee))
        .Cells(1, DispatchRegistryColumnAddressLine).value = CStr(firstItem(DispatchItemColumnAddressLine))
        .Cells(1, DispatchRegistryColumnEnvelopeFormatKey).value = CStr(firstItem(DispatchItemColumnEnvelopeFormatKey))
        .Cells(1, DispatchRegistryColumnMailType).value = CStr(firstItem(DispatchItemColumnMailType))
        .Cells(1, DispatchRegistryColumnMass).value = CStr(firstItem(DispatchItemColumnMass))
        .Cells(1, DispatchRegistryColumnDeclaredValue).value = CStr(firstItem(DispatchItemColumnDeclaredValue))
        .Cells(1, DispatchRegistryColumnPayment).value = ""
        .Cells(1, DispatchRegistryColumnComment).value = CStr(firstItem(DispatchItemColumnComment))
        .Cells(1, DispatchRegistryColumnPhone).value = CStr(firstItem(DispatchItemColumnPhone))
        .Cells(1, DispatchRegistryColumnIndexFrom).value = DispatchRepositoryGetSenderPostalCode(CStr(firstItem(DispatchItemColumnSenderName)))
        .Cells(1, DispatchRegistryColumnSenderName).value = CStr(firstItem(DispatchItemColumnSenderName))
        .Cells(1, DispatchRegistryColumnOutgoingNumbers).value = BuildBatchOutgoingNumbersText(batchItems)
        .Cells(1, DispatchRegistryColumnCreatedAt).value = Format$(Now, "dd.mm.yyyy hh:nn:ss")
        .Cells(1, DispatchRegistryColumnPostalCode).value = CStr(firstItem(DispatchItemColumnPostalCode))
    End With
End Sub

Private Sub UpdateLettersDispatchTracking(batchItems As Collection)
    Dim firstItem As Variant
    firstItem = batchItems(1)

    Dim batchId As String
    batchId = CStr(firstItem(DispatchItemColumnBatchId))

    Dim registryNumber As String
    registryNumber = CStr(firstItem(DispatchItemColumnRegistryNumber))

    Dim registryDate As String
    registryDate = CStr(firstItem(DispatchItemColumnRegistryDate))

    Dim i As Long
    For i = 1 To batchItems.count
        Dim dispatchItem As Variant
        dispatchItem = batchItems(i)

        Dim targetRowNumber As Long
        targetRowNumber = CLng(Val(CStr(dispatchItem(DispatchItemColumnLetterRowNumber))))

        If targetRowNumber < FIRST_DATA_ROW Then
            Call RepositoryTryResolveLetterRowNumber( _
                CStr(dispatchItem(DispatchItemColumnAddressee)), _
                CStr(dispatchItem(DispatchItemColumnLetterNumber)), _
                CStr(dispatchItem(DispatchItemColumnLetterDate)), _
                targetRowNumber)
        End If

        If targetRowNumber >= FIRST_DATA_ROW Then
            RepositoryUpdateLetterDispatchTracking _
                targetRowNumber, _
                t("history.dispatch_status.packed", "Да"), _
                batchId, _
                registryNumber, _
                registryDate
        End If
    Next i
End Sub

Private Function BuildBatchOutgoingNumbersText(batchItems As Collection) As String
    Dim parts As Collection
    Set parts = New Collection

    Dim i As Long
    For i = 1 To batchItems.count
        Dim dispatchItem As Variant
        dispatchItem = batchItems(i)
        AddCollectionTextPart parts, BuildDispatchOutgoingLine(dispatchItem)
    Next i

    BuildBatchOutgoingNumbersText = JoinCollectionWithDelimiter(parts, vbCrLf)
End Function

Private Function BuildDispatchOutgoingLine(dispatchItem As Variant) As String
    BuildDispatchOutgoingLine = Trim$(CStr(dispatchItem(DispatchItemColumnLetterNumber)))

    If Len(Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))) > 0 Then
        BuildDispatchOutgoingLine = BuildDispatchOutgoingLine & " " & t("common.preposition.from", "от") & " " & Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))
    End If
End Function

Private Sub AddCollectionTextPart(parts As Collection, textValue As String)
    If Len(Trim$(textValue)) = 0 Then Exit Sub
    parts.Add Trim$(textValue)
End Sub

Private Function JoinCollectionWithDelimiter(parts As Collection, delimiterText As String) As String
    Dim i As Long
    For i = 1 To parts.count
        If i > 1 Then
            JoinCollectionWithDelimiter = JoinCollectionWithDelimiter & delimiterText
        End If
        JoinCollectionWithDelimiter = JoinCollectionWithDelimiter & CStr(parts(i))
    Next i
End Function
