Attribute VB_Name = "ModuleDispatchJournal"
' ======================================================================
' Module: ModuleDispatchJournal
' Author: CreateLetter contributors
' Purpose: Build a dispatch package journal and safely return non-printed packages to work
' Version: 1.0.0 - 28.04.2026
' ======================================================================

Option Explicit

Private Const DispatchJournalSheetName As String = "DispatchJournal"
Private Const DispatchJournalTableName As String = "tblDispatchJournal"
Private lastReturnDispatchPackageMessage As String

Public Function BuildDispatchJournal() As Long
    On Error GoTo BuildError

    Dim targetSheet As Worksheet
    Set targetSheet = GetOrCreateDispatchJournalSheet()

    PrepareDispatchJournalSheet targetSheet
    WriteDispatchJournalHeaders targetSheet

    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()

    If dispatchItems Is Nothing Then
        CreateDispatchJournalTable targetSheet, 1
        targetSheet.Activate
        Exit Function
    End If

    If dispatchItems.count = 0 Then
        CreateDispatchJournalTable targetSheet, 1
        targetSheet.Activate
        Exit Function
    End If

    Dim groupedBatches As Object
    Set groupedBatches = GroupDispatchJournalItemsByBatch(dispatchItems)

    Dim batchKey As Variant
    Dim targetRow As Long
    targetRow = 2

    For Each batchKey In groupedBatches.keys
        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)

        If Not batchItems Is Nothing Then
            If batchItems.count > 0 Then
                WriteDispatchJournalRow targetSheet, targetRow, batchItems
                BuildDispatchJournal = BuildDispatchJournal + 1
                targetRow = targetRow + 1
            End If
        End If
    Next batchKey

    CreateDispatchJournalTable targetSheet, targetRow - 1
    FormatDispatchJournalSheet targetSheet, targetRow - 1
    targetSheet.Activate
    Exit Function

BuildError:
    Debug.Print "BuildDispatchJournal error: " & Err.description
    BuildDispatchJournal = 0
End Function

Public Sub OpenDispatchJournal()
    Dim packageCount As Long
    packageCount = BuildDispatchJournal()

    If packageCount = 0 Then
        MsgBox t("dispatch.journal.msg.no_items", "No dispatch packages found."), vbInformation, t("dispatch.journal.title", "Dispatch journal")
    End If
End Sub

Public Sub PromptReturnDispatchPackageToWork()
    On Error GoTo PromptError

    Dim batchId As String
    batchId = GetActiveDispatchPackageBatchId()

    If Len(Trim$(batchId)) = 0 Then
        batchId = InputBox(t("dispatch.journal.prompt.batch_id", "Select a package row in DispatchJournal or DispatchRegistry, or enter BatchId manually:"), t("dispatch.journal.return.title", "Return package to work"))
    End If

    batchId = Trim$(batchId)
    If Len(batchId) = 0 Then Exit Sub

    Dim confirmation As VbMsgBoxResult
    confirmation = MsgBox(t("dispatch.journal.return.confirm", "Return this package to available letters?") & vbCrLf & batchId, vbQuestion + vbYesNo, t("dispatch.journal.return.title", "Return package to work"))
    If confirmation <> vbYes Then Exit Sub

    If ReturnDispatchPackageToWork(batchId) Then
        MsgBox t("dispatch.journal.return.msg.done", "Package returned to available letters."), vbInformation, t("dispatch.journal.return.title", "Return package to work")
    Else
        If Len(Trim$(lastReturnDispatchPackageMessage)) = 0 Then lastReturnDispatchPackageMessage = t("dispatch.journal.return.msg.not_found", "Package was not found.")
        MsgBox lastReturnDispatchPackageMessage, vbExclamation, t("dispatch.journal.return.title", "Return package to work")
    End If

    Exit Sub

PromptError:
    MsgBox t("dispatch.journal.return.msg.error", "Failed to return package: ") & Err.description, vbCritical, t("dispatch.journal.return.title", "Return package to work")
End Sub

Public Function ReturnDispatchPackageToWork(ByVal batchId As String) As Boolean
    On Error GoTo ReturnError

    lastReturnDispatchPackageMessage = ""

    batchId = Trim$(batchId)
    If Len(batchId) = 0 Then Exit Function

    Dim dispatchTable As ListObject
    Set dispatchTable = GetDispatchItemsTable()

    If dispatchTable.DataBodyRange Is Nothing Then Exit Function

    Dim batchStatus As String
    batchStatus = GetDispatchBatchStatus(dispatchTable, batchId)

    If Len(batchStatus) = 0 Then
        lastReturnDispatchPackageMessage = t("dispatch.journal.return.msg.not_found", "Package was not found.")
        Exit Function
    End If

    If batchStatus = DispatchStatusRegistryPrinted Then
        lastReturnDispatchPackageMessage = t("dispatch.journal.return.error.printed", "Printed packages cannot be returned to work.")
        Exit Function
    End If

    Dim rowIndex As Long
    For rowIndex = dispatchTable.DataBodyRange.Rows.count To 1 Step -1
        If StrComp(CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnBatchId).value), batchId, vbTextCompare) = 0 Then
            ClearDispatchLetterTracking dispatchTable, rowIndex
            dispatchTable.DataBodyRange.Rows(rowIndex).Delete
            ReturnDispatchPackageToWork = True
        End If
    Next rowIndex

    If ReturnDispatchPackageToWork Then
        DeleteDispatchRegistryBatch batchId
        BuildDispatchJournal
    End If

    Exit Function

ReturnError:
    Debug.Print "ReturnDispatchPackageToWork error: " & Err.description
    Err.Raise Err.Number, "ReturnDispatchPackageToWork", Err.description
End Function

Private Function GetOrCreateDispatchJournalSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateDispatchJournalSheet = ThisWorkbook.Worksheets(DispatchJournalSheetName)
    On Error GoTo 0

    If GetOrCreateDispatchJournalSheet Is Nothing Then
        Set GetOrCreateDispatchJournalSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        GetOrCreateDispatchJournalSheet.Name = DispatchJournalSheetName
    End If

    GetOrCreateDispatchJournalSheet.Visible = xlSheetVisible
End Function

Private Sub PrepareDispatchJournalSheet(targetSheet As Worksheet)
    Dim tableIndex As Long
    For tableIndex = targetSheet.ListObjects.count To 1 Step -1
        targetSheet.ListObjects(tableIndex).Delete
    Next tableIndex

    targetSheet.Cells.Clear
    targetSheet.Cells.Font.Name = "Calibri"
    targetSheet.Cells.Font.Size = 11
End Sub

Private Sub WriteDispatchJournalHeaders(targetSheet As Worksheet)
    targetSheet.Cells(1, 1).value = t("dispatch.journal.column.batch_id", "Batch ID")
    targetSheet.Cells(1, 2).value = t("dispatch.journal.column.status", "Status")
    targetSheet.Cells(1, 3).value = t("dispatch.journal.column.registry_number", "Registry number")
    targetSheet.Cells(1, 4).value = t("dispatch.journal.column.registry_date", "Registry date")
    targetSheet.Cells(1, 5).value = t("dispatch.journal.column.addressee", "Addressee")
    targetSheet.Cells(1, 6).value = t("dispatch.journal.column.letter_count", "Letters")
    targetSheet.Cells(1, 7).value = t("dispatch.journal.column.outgoing_numbers", "Outgoing numbers")
    targetSheet.Cells(1, 8).value = t("dispatch.journal.column.sender", "Sender")
    targetSheet.Cells(1, 9).value = t("dispatch.journal.column.envelope_format", "Envelope")
    targetSheet.Cells(1, 10).value = t("dispatch.journal.column.mail_type", "Mail type")
    targetSheet.Cells(1, 11).value = t("dispatch.journal.column.created_at", "Created at")
    targetSheet.Cells(1, 12).value = t("dispatch.journal.column.comment", "Comment")
End Sub

Private Sub WriteDispatchJournalRow(targetSheet As Worksheet, ByVal targetRow As Long, batchItems As Collection)
    Dim firstItem As Variant
    firstItem = batchItems(1)

    targetSheet.Cells(targetRow, 1).value = CStr(firstItem(DispatchItemColumnBatchId))
    targetSheet.Cells(targetRow, 2).value = FormatDispatchJournalStatus(GetBatchItemsStatus(batchItems))
    targetSheet.Cells(targetRow, 3).value = CStr(firstItem(DispatchItemColumnRegistryNumber))
    targetSheet.Cells(targetRow, 4).value = CStr(firstItem(DispatchItemColumnRegistryDate))
    targetSheet.Cells(targetRow, 5).value = CStr(firstItem(DispatchItemColumnAddressee))
    targetSheet.Cells(targetRow, 6).value = batchItems.count
    targetSheet.Cells(targetRow, 7).value = BuildDispatchJournalOutgoingNumbers(batchItems)
    targetSheet.Cells(targetRow, 8).value = CStr(firstItem(DispatchItemColumnSenderName))
    targetSheet.Cells(targetRow, 9).value = UCase$(CStr(firstItem(DispatchItemColumnEnvelopeFormatKey)))
    targetSheet.Cells(targetRow, 10).value = CStr(firstItem(DispatchItemColumnMailType))
    targetSheet.Cells(targetRow, 11).value = CStr(firstItem(DispatchItemColumnCreatedAt))
    targetSheet.Cells(targetRow, 12).value = CStr(firstItem(DispatchItemColumnComment))
End Sub

Private Sub CreateDispatchJournalTable(targetSheet As Worksheet, ByVal lastRow As Long)
    If lastRow < 1 Then lastRow = 1

    Dim sourceRange As Range
    Set sourceRange = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(lastRow, 12))

    Dim journalTable As ListObject
    Set journalTable = targetSheet.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
    journalTable.Name = DispatchJournalTableName
    journalTable.TableStyle = "TableStyleMedium2"
End Sub

Private Sub FormatDispatchJournalSheet(targetSheet As Worksheet, ByVal lastRow As Long)
    targetSheet.Columns("A").ColumnWidth = 36
    targetSheet.Columns("B").ColumnWidth = 20
    targetSheet.Columns("C").ColumnWidth = 16
    targetSheet.Columns("D").ColumnWidth = 14
    targetSheet.Columns("E").ColumnWidth = 30
    targetSheet.Columns("F").ColumnWidth = 10
    targetSheet.Columns("G").ColumnWidth = 28
    targetSheet.Columns("H").ColumnWidth = 18
    targetSheet.Columns("I").ColumnWidth = 12
    targetSheet.Columns("J").ColumnWidth = 18
    targetSheet.Columns("K").ColumnWidth = 20
    targetSheet.Columns("L").ColumnWidth = 28

    If lastRow >= 2 Then
        targetSheet.Range(targetSheet.Cells(2, 7), targetSheet.Cells(lastRow, 7)).WrapText = True
        targetSheet.Range(targetSheet.Cells(2, 1), targetSheet.Cells(lastRow, 12)).VerticalAlignment = xlTop
    End If
End Sub

Private Function GroupDispatchJournalItemsByBatch(dispatchItems As Collection) As Object
    Dim groupedBatches As Object
    Set groupedBatches = CreateObject("Scripting.Dictionary")
    groupedBatches.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To dispatchItems.count
        Dim dispatchItem As Variant
        dispatchItem = dispatchItems(i)

        Dim batchKey As String
        batchKey = Trim$(CStr(dispatchItem(DispatchItemColumnBatchId)))
        If Len(batchKey) = 0 Then batchKey = CStr(dispatchItem(DispatchItemColumnId))

        If Not groupedBatches.Exists(batchKey) Then groupedBatches.Add batchKey, New Collection

        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)
        batchItems.Add dispatchItem
    Next i

    Set GroupDispatchJournalItemsByBatch = groupedBatches
End Function

Private Function GetBatchItemsStatus(batchItems As Collection) As String
    Dim resultStatus As String
    resultStatus = DispatchStatusDraft

    Dim i As Long
    For i = 1 To batchItems.count
        Dim dispatchItem As Variant
        dispatchItem = batchItems(i)
        resultStatus = PickHigherDispatchStatus(resultStatus, LCase$(Trim$(CStr(dispatchItem(DispatchItemColumnStatus)))))
    Next i

    GetBatchItemsStatus = resultStatus
End Function

Private Function GetDispatchBatchStatus(dispatchTable As ListObject, ByVal batchId As String) As String
    Dim rowIndex As Long
    For rowIndex = 1 To dispatchTable.DataBodyRange.Rows.count
        If StrComp(CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnBatchId).value), batchId, vbTextCompare) = 0 Then
            GetDispatchBatchStatus = PickHigherDispatchStatus(GetDispatchBatchStatus, LCase$(Trim$(CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnStatus).value))))
        End If
    Next rowIndex
End Function

Private Function PickHigherDispatchStatus(ByVal currentStatus As String, ByVal candidateStatus As String) As String
    If Len(currentStatus) = 0 Then currentStatus = DispatchStatusDraft
    If Len(candidateStatus) = 0 Then candidateStatus = DispatchStatusDraft

    If candidateStatus = DispatchStatusRegistryPrinted Or currentStatus = DispatchStatusRegistryPrinted Then
        PickHigherDispatchStatus = DispatchStatusRegistryPrinted
    ElseIf candidateStatus = DispatchStatusRegistered Or currentStatus = DispatchStatusRegistered Then
        PickHigherDispatchStatus = DispatchStatusRegistered
    ElseIf candidateStatus = DispatchStatusPacked Or currentStatus = DispatchStatusPacked Then
        PickHigherDispatchStatus = DispatchStatusPacked
    Else
        PickHigherDispatchStatus = currentStatus
    End If
End Function

Private Function FormatDispatchJournalStatus(ByVal status As String) As String
    Select Case LCase$(Trim$(status))
    Case DispatchStatusPacked
        FormatDispatchJournalStatus = t("dispatch.journal.status.packed", "Packed")
    Case DispatchStatusRegistered
        FormatDispatchJournalStatus = t("dispatch.journal.status.registered", "Registered")
    Case DispatchStatusRegistryPrinted
        FormatDispatchJournalStatus = t("dispatch.journal.status.registry_printed", "Printed")
    Case Else
        FormatDispatchJournalStatus = t("dispatch.journal.status.draft", "Draft")
    End Select
End Function

Private Function BuildDispatchJournalOutgoingNumbers(batchItems As Collection) As String
    Dim i As Long
    For i = 1 To batchItems.count
        Dim dispatchItem As Variant
        dispatchItem = batchItems(i)

        If Len(BuildDispatchJournalOutgoingNumbers) > 0 Then BuildDispatchJournalOutgoingNumbers = BuildDispatchJournalOutgoingNumbers & vbCrLf

        BuildDispatchJournalOutgoingNumbers = BuildDispatchJournalOutgoingNumbers & CStr(dispatchItem(DispatchItemColumnLetterNumber))

        If Len(Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))) > 0 Then
            BuildDispatchJournalOutgoingNumbers = BuildDispatchJournalOutgoingNumbers & " " & t("common.preposition.from", "dated") & " " & CStr(dispatchItem(DispatchItemColumnLetterDate))
        End If
    Next i
End Function

Private Sub ClearDispatchLetterTracking(dispatchTable As ListObject, ByVal rowIndex As Long)
    Dim targetRowNumber As Long
    targetRowNumber = CLng(Val(CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnLetterRowNumber).value)))

    If targetRowNumber < FIRST_DATA_ROW Then
        RepositoryTryResolveLetterRowNumber CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnAddressee).value), CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnLetterNumber).value), CStr(dispatchTable.DataBodyRange.Cells(rowIndex, DispatchItemColumnLetterDate).value), targetRowNumber
    End If

    If targetRowNumber >= FIRST_DATA_ROW Then RepositoryUpdateLetterDispatchTracking targetRowNumber, "", "", "", ""
End Sub

Private Sub DeleteDispatchRegistryBatch(ByVal batchId As String)
    On Error GoTo DeleteError

    Dim registryTable As ListObject
    Set registryTable = GetDispatchRegistryTable()

    If registryTable.DataBodyRange Is Nothing Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = registryTable.DataBodyRange.Rows.count To 1 Step -1
        If StrComp(CStr(registryTable.DataBodyRange.Cells(rowIndex, DispatchRegistryColumnBatchId).value), batchId, vbTextCompare) = 0 Then
            registryTable.DataBodyRange.Rows(rowIndex).Delete
        End If
    Next rowIndex

    Exit Sub

DeleteError:
    Debug.Print "DeleteDispatchRegistryBatch error: " & Err.description
End Sub

Private Function GetActiveDispatchPackageBatchId() As String
    On Error GoTo ResolveError

    If ActiveSheet Is Nothing Then Exit Function
    If ActiveCell Is Nothing Then Exit Function
    If ActiveCell.Row < 2 Then Exit Function

    Select Case LCase$(ActiveSheet.Name)
    Case LCase$(DispatchJournalSheetName)
        GetActiveDispatchPackageBatchId = Trim$(CStr(ActiveSheet.Cells(ActiveCell.Row, 1).value))
    Case "dispatchregistry"
        GetActiveDispatchPackageBatchId = Trim$(CStr(ActiveSheet.Cells(ActiveCell.Row, DispatchRegistryColumnBatchId).value))
    Case "dispatchitems"
        GetActiveDispatchPackageBatchId = Trim$(CStr(ActiveSheet.Cells(ActiveCell.Row, DispatchItemColumnBatchId).value))
    End Select

    Exit Function

ResolveError:
    GetActiveDispatchPackageBatchId = ""
End Function

Private Function GetDispatchItemsTable() As ListObject
    Set GetDispatchItemsTable = ThisWorkbook.Worksheets("DispatchItems").ListObjects.item(DispatchItemsTableName)
End Function

Private Function GetDispatchRegistryTable() As ListObject
    Set GetDispatchRegistryTable = ThisWorkbook.Worksheets("DispatchRegistry").ListObjects.item(DispatchRegistryTableName)
End Function
