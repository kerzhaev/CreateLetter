Attribute VB_Name = "ModuleDispatchRegistry"
' ======================================================================
' Module: ModuleDispatchRegistry
' Author: CreateLetter contributors
' Purpose: Build and refresh the internal Excel dispatch registry from grouped dispatch packages
' Version: 1.2.3 - 01.05.2026
' ======================================================================

Option Explicit

Public Function BuildDispatchRegistryFromDispatchItems() As Long
    On Error GoTo BuildError

    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()

    If dispatchItems Is Nothing Or dispatchItems.count = 0 Then Exit Function

    Dim registryScopeItems As Collection
    Set registryScopeItems = FilterDispatchItemsForNextRegistry(dispatchItems)
    If registryScopeItems Is Nothing Or registryScopeItems.count = 0 Then Exit Function

    ClearDispatchRegistry

    Dim groupedBatches As Object
    Set groupedBatches = GroupDispatchItemsByBatch(registryScopeItems)
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
                Dim firstBatchItem As Variant
                firstBatchItem = batchItems(1)
                AppendDispatchRegistryRow registryTable, batchItems
                UpdateLettersDispatchTracking batchItems
                DispatchRepositoryUpdateBatchRegistryState CStr(batchKey), CStr(firstBatchItem(DispatchItemColumnRegistryNumber)), CStr(firstBatchItem(DispatchItemColumnRegistryDate)), DispatchStatusRegistered
                BuildDispatchRegistryFromDispatchItems = BuildDispatchRegistryFromDispatchItems + 1
            End If
        End If
    Next batchKey

    Exit Function

BuildError:
    Debug.Print "BuildDispatchRegistryFromDispatchItems error: " & Err.description
    BuildDispatchRegistryFromDispatchItems = 0
End Function

Private Function FilterDispatchItemsForNextRegistry(dispatchItems As Collection) As Collection
    Set FilterDispatchItemsForNextRegistry = New Collection

    Dim targetRegistryKeys As Object
    Set targetRegistryKeys = CreateObject("Scripting.Dictionary")
    targetRegistryKeys.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To dispatchItems.count
        Dim candidateItem As Variant
        candidateItem = dispatchItems(i)

        If IsDispatchItemReadyForRegistryBuild(candidateItem) Then
            Dim targetKey As String
            targetKey = BuildDispatchItemRegistryKey(candidateItem)

            If Len(targetKey) > 0 Then
                If Not targetRegistryKeys.Exists(targetKey) Then targetRegistryKeys.Add targetKey, True
            End If
        End If
    Next i

    If targetRegistryKeys.count = 0 Then Exit Function

    For i = 1 To dispatchItems.count
        Dim dispatchItem As Variant
        dispatchItem = dispatchItems(i)

        If DispatchItemBelongsToRegistryBuildScope(dispatchItem, targetRegistryKeys) Then
            FilterDispatchItemsForNextRegistry.Add dispatchItem
        End If
    Next i
End Function

Private Function IsDispatchItemReadyForRegistryBuild(dispatchItem As Variant) As Boolean
    IsDispatchItemReadyForRegistryBuild = LCase$(Trim$(CStr(dispatchItem(DispatchItemColumnStatus)))) = DispatchStatusPacked
End Function

Private Function DispatchItemBelongsToRegistryBuildScope(dispatchItem As Variant, targetRegistryKeys As Object) As Boolean
    Dim registryKey As String
    registryKey = BuildDispatchItemRegistryKey(dispatchItem)

    If Len(registryKey) = 0 Then Exit Function
    If Not targetRegistryKeys.Exists(registryKey) Then Exit Function

    Dim currentStatus As String
    currentStatus = LCase$(Trim$(CStr(dispatchItem(DispatchItemColumnStatus))))

    DispatchItemBelongsToRegistryBuildScope = currentStatus = DispatchStatusPacked Or currentStatus = DispatchStatusRegistered
End Function

Private Function BuildDispatchItemRegistryKey(dispatchItem As Variant) As String
    Dim registryNumber As String
    registryNumber = Trim$(CStr(dispatchItem(DispatchItemColumnRegistryNumber)))

    Dim registryDate As String
    registryDate = Trim$(CStr(dispatchItem(DispatchItemColumnRegistryDate)))

    If Len(registryNumber) = 0 Then Exit Function
    If Len(registryDate) = 0 Then Exit Function

    BuildDispatchItemRegistryKey = registryNumber & "|" & registryDate
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

Public Function CountDispatchRegistryRows() As Long
    On Error GoTo CountError

    Dim registryTable As ListObject
    Set registryTable = GetDispatchRegistryTable()

    If registryTable.DataBodyRange Is Nothing Then Exit Function

    CountDispatchRegistryRows = registryTable.DataBodyRange.Rows.count
    Exit Function

CountError:
    Debug.Print "CountDispatchRegistryRows error: " & Err.description
    CountDispatchRegistryRows = 0
End Function

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
