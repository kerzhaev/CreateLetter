Attribute VB_Name = "ModuleEnvelopeLayouts"
' ======================================================================
' Module: ModuleEnvelopeLayouts
' Author: CreateLetter contributors
' Purpose: Prepare hidden workbook layout sheets for grouped C4, C5, and DL envelope batches
' Version: 1.1.0 - 26.04.2026
' ======================================================================

Option Explicit

Private Const EnvelopeLayoutSheetC4 As String = "DispatchLayout_C4"
Private Const EnvelopeLayoutSheetC5 As String = "DispatchLayout_C5"
Private Const EnvelopeLayoutSheetDL As String = "DispatchLayout_DL"

Public Function PrepareEnvelopePrint() As Long
    On Error GoTo PrepareError

    EnsureEnvelopeLayoutSheets
    ClearEnvelopeLayoutData

    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()
    If dispatchItems Is Nothing Or dispatchItems.count = 0 Then Exit Function

    Dim groupedBatches As Object
    Set groupedBatches = GroupDispatchItemsByBatch(dispatchItems)
    If groupedBatches Is Nothing Then Exit Function

    Dim batchKey As Variant
    For Each batchKey In groupedBatches.keys
        Dim batchItems As Collection
        Set batchItems = groupedBatches(batchKey)
        If Not batchItems Is Nothing Then
            If batchItems.count > 0 Then
                If AppendEnvelopeLayoutRow(batchItems) Then
                    PrepareEnvelopePrint = PrepareEnvelopePrint + 1
                End If
            End If
        End If
    Next batchKey

    Exit Function

PrepareError:
    Debug.Print "PrepareEnvelopePrint error: " & Err.description
    PrepareEnvelopePrint = 0
End Function

Public Function ResolveEnvelopeLayoutSheetName(envelopeFormatKey As String) As String
    Select Case LCase$(Trim$(envelopeFormatKey))
    Case "c4"
        ResolveEnvelopeLayoutSheetName = EnvelopeLayoutSheetC4
    Case "c5"
        ResolveEnvelopeLayoutSheetName = EnvelopeLayoutSheetC5
    Case "dl"
        ResolveEnvelopeLayoutSheetName = EnvelopeLayoutSheetDL
    End Select
End Function

Private Sub EnsureEnvelopeLayoutSheets()
    EnsureEnvelopeLayoutSheet EnvelopeLayoutSheetC4
    EnsureEnvelopeLayoutSheet EnvelopeLayoutSheetC5
    EnsureEnvelopeLayoutSheet EnvelopeLayoutSheetDL
End Sub

Private Sub EnsureEnvelopeLayoutSheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub ClearEnvelopeLayoutData()
    ClearEnvelopeLayoutSheetData EnvelopeLayoutSheetC4
    ClearEnvelopeLayoutSheetData EnvelopeLayoutSheetC5
    ClearEnvelopeLayoutSheetData EnvelopeLayoutSheetDL
End Sub

Private Sub ClearEnvelopeLayoutSheetData(sheetName As String)
    On Error GoTo ClearError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 15)).ClearContents
    End If

    Exit Sub

ClearError:
    Debug.Print "ClearEnvelopeLayoutSheetData error: " & Err.description
End Sub

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

Private Function AppendEnvelopeLayoutRow(batchItems As Collection) As Boolean
    Dim firstItem As Variant
    firstItem = batchItems(1)

    Dim sheetName As String
    sheetName = ResolveEnvelopeLayoutSheetName(CStr(firstItem(DispatchItemColumnEnvelopeFormatKey)))
    If Len(sheetName) = 0 Then Exit Function

    On Error GoTo AppendError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim NextRow As Long
    NextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2

    ws.Cells(NextRow, 1).value = CStr(firstItem(DispatchItemColumnBatchId))
    ws.Cells(NextRow, 2).value = CStr(firstItem(DispatchItemColumnRegistryNumber))
    ws.Cells(NextRow, 3).value = CStr(firstItem(DispatchItemColumnRegistryDate))
    ws.Cells(NextRow, 4).value = CStr(firstItem(DispatchItemColumnAddressee))
    ws.Cells(NextRow, 5).value = CStr(firstItem(DispatchItemColumnAddressLine))
    ws.Cells(NextRow, 6).value = CStr(firstItem(DispatchItemColumnPostalCode))
    ws.Cells(NextRow, 7).value = CStr(firstItem(DispatchItemColumnSenderName))
    ws.Cells(NextRow, 8).value = DispatchRepositoryGetSenderPostalCode(CStr(firstItem(DispatchItemColumnSenderName)))
    ws.Cells(NextRow, 9).value = BuildBatchOutgoingNumbersText(batchItems)
    ws.Cells(NextRow, 10).value = CStr(firstItem(DispatchItemColumnEnvelopeFormatKey))
    ws.Cells(NextRow, 11).value = CStr(firstItem(DispatchItemColumnMailType))
    ws.Cells(NextRow, 12).value = CStr(firstItem(DispatchItemColumnMass))
    ws.Cells(NextRow, 13).value = CStr(firstItem(DispatchItemColumnDeclaredValue))
    ws.Cells(NextRow, 14).value = CStr(firstItem(DispatchItemColumnComment))
    ws.Cells(NextRow, 15).value = Format$(Now, "dd.mm.yyyy hh:nn:ss")

    AppendEnvelopeLayoutRow = True
    Exit Function

AppendError:
    Debug.Print "AppendEnvelopeLayoutRow error: " & Err.description
    AppendEnvelopeLayoutRow = False
End Function

Private Function BuildBatchOutgoingNumbersText(batchItems As Collection) As String
    Dim i As Long
    For i = 1 To batchItems.count
        Dim dispatchItem As Variant
        dispatchItem = batchItems(i)

        If i > 1 Then
            BuildBatchOutgoingNumbersText = BuildBatchOutgoingNumbersText & vbCrLf
        End If

        BuildBatchOutgoingNumbersText = BuildBatchOutgoingNumbersText & Trim$(CStr(dispatchItem(DispatchItemColumnLetterNumber)))

        If Len(Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))) > 0 Then
            BuildBatchOutgoingNumbersText = BuildBatchOutgoingNumbersText & " " & t("common.preposition.from", "от") & " " & Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))
        End If
    Next i
End Function