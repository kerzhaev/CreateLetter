Attribute VB_Name = "ModuleDispatchRegistry"
' ======================================================================
' Module: ModuleDispatchRegistry
' Author: CreateLetter contributors
' Purpose: Build and refresh the internal Excel dispatch registry from DispatchItems
' Version: 1.0.0 - 26.04.2026
' ======================================================================

Option Explicit

Public Function BuildDispatchRegistryFromDispatchItems() As Long
    On Error GoTo BuildError

    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()

    ClearDispatchRegistry

    If dispatchItems Is Nothing Or dispatchItems.count = 0 Then Exit Function

    Dim registryTable As ListObject
    Set registryTable = GetDispatchRegistryTable()

    Dim i As Long
    For i = 1 To dispatchItems.count
        AppendDispatchRegistryRow registryTable, dispatchItems(i)
        BuildDispatchRegistryFromDispatchItems = BuildDispatchRegistryFromDispatchItems + 1
    Next i

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
    Set GetDispatchRegistryTable = ws.ListObjects.Item(DispatchRegistryTableName)
End Function

Private Sub AppendDispatchRegistryRow(registryTable As ListObject, dispatchItem As Variant)
    Dim newRow As ListRow
    Set newRow = registryTable.ListRows.Add

    With newRow.Range
        .Cells(1, DispatchRegistryColumnAddressLine).value = CStr(dispatchItem(DispatchItemColumnAddressLine))
        .Cells(1, DispatchRegistryColumnAddressee).value = CStr(dispatchItem(DispatchItemColumnAddressee))
        .Cells(1, DispatchRegistryColumnMailType).value = CStr(dispatchItem(DispatchItemColumnMailType))
        .Cells(1, DispatchRegistryColumnEnvelopeFormatKey).value = CStr(dispatchItem(DispatchItemColumnEnvelopeFormatKey))
        .Cells(1, DispatchRegistryColumnMass).value = CStr(dispatchItem(DispatchItemColumnMass))
        .Cells(1, DispatchRegistryColumnDeclaredValue).value = CStr(dispatchItem(DispatchItemColumnDeclaredValue))
        .Cells(1, DispatchRegistryColumnPayment).value = ""
        .Cells(1, DispatchRegistryColumnComment).value = CStr(dispatchItem(DispatchItemColumnComment))
        .Cells(1, DispatchRegistryColumnPhone).value = CStr(dispatchItem(DispatchItemColumnPhone))
        .Cells(1, DispatchRegistryColumnIndexFrom).value = DispatchRepositoryGetSenderPostalCode(CStr(dispatchItem(DispatchItemColumnSenderName)))
        .Cells(1, DispatchRegistryColumnBatchId).value = CStr(dispatchItem(DispatchItemColumnBatchId))
        .Cells(1, DispatchRegistryColumnCreatedAt).value = CStr(dispatchItem(DispatchItemColumnCreatedAt))
    End With
End Sub
