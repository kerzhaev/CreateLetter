Attribute VB_Name = "ModuleEnvelopeLayouts"
' ======================================================================
' Module: ModuleEnvelopeLayouts
' Author: CreateLetter contributors
' Purpose: Prepare printable workbook layout sheets for grouped C4, C5, and DL envelope batches
' Version: 1.5.5 - 02.05.2026
' ======================================================================

Option Explicit

Private Const EnvelopeLayoutSheetC4 As String = "DispatchLayout_C4"
Private Const EnvelopeLayoutSheetC5 As String = "DispatchLayout_C5"
Private Const EnvelopeLayoutSheetDL As String = "DispatchLayout_DL"
Private Const EnvelopeLayoutFirstColumn As Long = 1
Private Const EnvelopeLayoutLastColumn As Long = 6
Private Const EnvelopeLayoutLastColumnC4 As Long = 11
Private Const EnvelopeLayoutLastColumnTemplate As Long = 11
Private Const EnvelopePreviewShapePrefix As String = "EnvelopePreviewGrid_"
Private Const EnvelopeDynamicShapePrefix As String = "EnvelopeDynamic_"

Public Function PrepareEnvelopePrint() As Long
    PrepareEnvelopePrint = PrepareEnvelopeLayoutsForBatch(False, "")
End Function

Public Function PrepareEnvelopePreviewGrid() As Long
    PrepareEnvelopePreviewGrid = PrepareEnvelopeLayoutsForBatch(True, "")
End Function

Public Function PrepareEnvelopePrintForBatch(batchId As String) As Long
    PrepareEnvelopePrintForBatch = PrepareEnvelopeLayoutsForBatch(False, batchId)
End Function

Public Sub PreviewPreparedEnvelopeForBatch(batchId As String)
    On Error GoTo PreviewError

    Dim sheetName As String
    sheetName = GetEnvelopeSheetNameForBatch(batchId)
    If Len(sheetName) = 0 Then Exit Sub

    If PrepareEnvelopePrintForBatch(batchId) <= 0 Then Exit Sub

    ThisWorkbook.Worksheets(sheetName).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(sheetName).Activate
    ThisWorkbook.Worksheets(sheetName).PrintPreview
    ThisWorkbook.Worksheets(sheetName).Visible = xlSheetVeryHidden

    Exit Sub

PreviewError:
    Debug.Print "PreviewPreparedEnvelopeForBatch error: " & Err.description
End Sub

Private Function PrepareEnvelopeLayoutsForBatch(showPreviewGrid As Boolean, targetBatchId As String) As Long
    On Error GoTo PrepareError

    EnsureEnvelopeLayoutSheets
    ClearEnvelopeLayoutData

    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()
    If dispatchItems Is Nothing Or dispatchItems.count = 0 Then Exit Function

    Dim registryBatchIds As Object
    Set registryBatchIds = GetCurrentRegistryBatchIdSet()
    If registryBatchIds Is Nothing Then Exit Function
    If registryBatchIds.count = 0 Then Exit Function

    If Len(Trim$(targetBatchId)) > 0 Then Set registryBatchIds = BuildSingleEnvelopeBatchIdSet(registryBatchIds, targetBatchId)
    If registryBatchIds Is Nothing Then Exit Function
    If registryBatchIds.count = 0 Then Exit Function

    Dim scopedDispatchItems As Collection
    Set scopedDispatchItems = FilterDispatchItemsByBatchIdSet(dispatchItems, registryBatchIds)
    If scopedDispatchItems Is Nothing Or scopedDispatchItems.count = 0 Then Exit Function

    Dim groupedBatches As Object
    Set groupedBatches = GroupDispatchItemsByBatch(scopedDispatchItems)
    If groupedBatches Is Nothing Then Exit Function

    Dim firstVisibleSheet As Worksheet
    Dim batchKey As Variant
    For Each batchKey In groupedBatches.keys
        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)
        If Not batchItems Is Nothing Then
            If batchItems.count > 0 Then
                If AppendEnvelopeLayoutPage(batchItems, firstVisibleSheet, showPreviewGrid) Then
                    PrepareEnvelopeLayoutsForBatch = PrepareEnvelopeLayoutsForBatch + 1
                End If
            End If
        End If
    Next batchKey

    FinalizeEnvelopeLayoutSheet EnvelopeLayoutSheetC4
    FinalizeEnvelopeLayoutSheet EnvelopeLayoutSheetC5
    FinalizeEnvelopeLayoutSheet EnvelopeLayoutSheetDL

    If Not firstVisibleSheet Is Nothing Then firstVisibleSheet.Activate
    Exit Function

PrepareError:
    Debug.Print "PrepareEnvelopeLayoutsForBatch error: " & Err.description
    PrepareEnvelopeLayoutsForBatch = 0
End Function

Private Function GetEnvelopeSheetNameForBatch(batchId As String) As String
    Dim dispatchItems As Collection
    Set dispatchItems = DispatchRepositoryLoadDispatchItems()

    If dispatchItems Is Nothing Then Exit Function

    Dim normalizedBatchId As String
    normalizedBatchId = Trim$(batchId)
    If Len(normalizedBatchId) = 0 Then Exit Function

    Dim itemIndex As Long
    For itemIndex = 1 To dispatchItems.count
        Dim dispatchItem As Variant
        dispatchItem = dispatchItems(itemIndex)

        If StrComp(Trim$(CStr(dispatchItem(DispatchItemColumnBatchId))), normalizedBatchId, vbTextCompare) = 0 Then
            GetEnvelopeSheetNameForBatch = ResolveEnvelopeLayoutSheetName(CStr(dispatchItem(DispatchItemColumnEnvelopeFormatKey)))
            Exit Function
        End If
    Next itemIndex
End Function

Private Function BuildSingleEnvelopeBatchIdSet(registryBatchIds As Object, targetBatchId As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim normalizedBatchId As String
    normalizedBatchId = Trim$(targetBatchId)

    If Len(normalizedBatchId) > 0 Then
        If Not registryBatchIds Is Nothing Then
            If registryBatchIds.Exists(normalizedBatchId) Then result.Add normalizedBatchId, True
        End If
    End If

    Set BuildSingleEnvelopeBatchIdSet = result
End Function

Private Function GetCurrentRegistryBatchIdSet() As Object
    On Error GoTo RegistryError

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DispatchRegistry")

    Dim registryTable As ListObject
    Set registryTable = ws.ListObjects.item(DispatchRegistryTableName)
    If registryTable.DataBodyRange Is Nothing Then
        Set GetCurrentRegistryBatchIdSet = result
        Exit Function
    End If

    Dim rowIndex As Long
    For rowIndex = 1 To registryTable.DataBodyRange.Rows.count
        Dim batchId As String
        batchId = Trim$(CStr(registryTable.DataBodyRange.Cells(rowIndex, DispatchRegistryColumnBatchId).value))
        If Len(batchId) > 0 Then result.item(batchId) = True
    Next rowIndex

    Set GetCurrentRegistryBatchIdSet = result
    Exit Function

RegistryError:
    Debug.Print "GetCurrentRegistryBatchIdSet error: " & Err.description
    Set GetCurrentRegistryBatchIdSet = Nothing
End Function

Private Function FilterDispatchItemsByBatchIdSet(dispatchItems As Collection, registryBatchIds As Object) As Collection
    Dim result As Collection
    Set result = New Collection

    If dispatchItems Is Nothing Then
        Set FilterDispatchItemsByBatchIdSet = result
        Exit Function
    End If

    If registryBatchIds Is Nothing Then
        Set FilterDispatchItemsByBatchIdSet = result
        Exit Function
    End If

    Dim itemIndex As Long
    For itemIndex = 1 To dispatchItems.count
        Dim dispatchItem As Variant
        dispatchItem = dispatchItems(itemIndex)

        Dim batchId As String
        batchId = Trim$(CStr(dispatchItem(DispatchItemColumnBatchId)))

        If Len(batchId) > 0 Then
            If registryBatchIds.Exists(batchId) Then result.Add dispatchItem
        End If
    Next itemIndex

    Set FilterDispatchItemsByBatchIdSet = result
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

    ws.Visible = xlSheetVisible
    ClearEnvelopePreviewShapes ws
    ws.ResetAllPageBreaks

    Dim envelopeFormatKey As String
    envelopeFormatKey = ResolveEnvelopeFormatKeyFromSheetName(sheetName)

    Dim rowsPerPage As Long
    rowsPerPage = GetEnvelopeRowsPerPage(envelopeFormatKey)

    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1

    If lastRow > rowsPerPage Then
        ws.Range(CStr(rowsPerPage + 1) & ":" & CStr(lastRow)).UnMerge
        ws.Range(CStr(rowsPerPage + 1) & ":" & CStr(lastRow)).Clear
    End If

    ClearEnvelopeTemplateDynamicCells ws, envelopeFormatKey, 1
    ws.Visible = xlSheetVeryHidden

    Exit Sub

ClearError:
    Debug.Print "ClearEnvelopeLayoutSheetData error: " & Err.description
End Sub

Private Sub ClearEnvelopePreviewShapes(ws As Worksheet)
    On Error GoTo ClearShapesError

    Dim shapeIndex As Long
    For shapeIndex = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes.item(shapeIndex).Name, Len(EnvelopePreviewShapePrefix)) = EnvelopePreviewShapePrefix Then ws.Shapes.item(shapeIndex).Delete
        If Left$(ws.Shapes.item(shapeIndex).Name, Len(EnvelopeDynamicShapePrefix)) = EnvelopeDynamicShapePrefix Then ws.Shapes.item(shapeIndex).Delete
    Next shapeIndex

    Exit Sub

ClearShapesError:
    Debug.Print "ClearEnvelopePreviewShapes error: " & Err.description
End Sub

Private Sub ClearEnvelopeRuntimeShapesInPage(ws As Worksheet, topRow As Long, rowsPerPage As Long)
    On Error GoTo ClearPageShapesError

    Dim pageTop As Double
    pageTop = ws.Rows(topRow).Top - 2

    Dim pageBottom As Double
    pageBottom = ws.Rows(topRow + rowsPerPage - 1).Top + ws.Rows(topRow + rowsPerPage - 1).Height + 2

    Dim shapeIndex As Long
    For shapeIndex = ws.Shapes.count To 1 Step -1
        If IsEnvelopeRuntimeShapeName(ws.Shapes.item(shapeIndex).Name) Then
            If ws.Shapes.item(shapeIndex).Top >= pageTop And ws.Shapes.item(shapeIndex).Top <= pageBottom Then ws.Shapes.item(shapeIndex).Delete
        End If
    Next shapeIndex

    Exit Sub

ClearPageShapesError:
    Debug.Print "ClearEnvelopeRuntimeShapesInPage error: " & Err.description
End Sub

Private Function IsEnvelopeRuntimeShapeName(shapeName As String) As Boolean
    If Left$(shapeName, Len(EnvelopePreviewShapePrefix)) = EnvelopePreviewShapePrefix Then IsEnvelopeRuntimeShapeName = True
    If Left$(shapeName, Len(EnvelopeDynamicShapePrefix)) = EnvelopeDynamicShapePrefix Then IsEnvelopeRuntimeShapeName = True
End Function

Private Function ResolveEnvelopeFormatKeyFromSheetName(sheetName As String) As String
    Select Case sheetName
    Case EnvelopeLayoutSheetC4
        ResolveEnvelopeFormatKeyFromSheetName = "c4"
    Case EnvelopeLayoutSheetC5
        ResolveEnvelopeFormatKeyFromSheetName = "c5"
    Case EnvelopeLayoutSheetDL
        ResolveEnvelopeFormatKeyFromSheetName = "dl"
    Case Else
        ResolveEnvelopeFormatKeyFromSheetName = "c4"
    End Select
End Function

Private Sub ClearEnvelopeTemplateDynamicCells(ws As Worksheet, envelopeFormatKey As String, topRow As Long)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetMailTypeMarkCell(envelopeFormatKey, 1)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetMailTypeMarkCell(envelopeFormatKey, 2)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetSenderNameCell(envelopeFormatKey)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetSenderAddressCell(envelopeFormatKey, 1)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetSenderAddressCell(envelopeFormatKey, 2)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetSenderAddressCell(envelopeFormatKey, 3)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetSenderPostalCell(envelopeFormatKey)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetOutgoingCell(envelopeFormatKey)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetRecipientNameCell(envelopeFormatKey)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetRecipientAddressCell(envelopeFormatKey, 1)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetRecipientAddressCell(envelopeFormatKey, 2)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetRecipientAddressCell(envelopeFormatKey, 3)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetRecipientPostalCell(envelopeFormatKey)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetLargePostalCell(envelopeFormatKey)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetBatchMarkerCell(envelopeFormatKey)
End Sub

Private Sub ClearTemplateCell(ws As Worksheet, envelopeFormatKey As String, topRow As Long, cellAddress As String)
    If Len(cellAddress) = 0 Then Exit Sub

    Dim targetCell As Range
    Set targetCell = ws.Range(GetOffsetCellAddress(ws, cellAddress, topRow))

    If targetCell.MergeCells Then
        targetCell.MergeArea.ClearContents
    Else
        targetCell.ClearContents
    End If
End Sub

Private Sub SetTemplateValue(ws As Worksheet, topRow As Long, cellAddress As String, cellValue As String)
    If Len(cellAddress) = 0 Then Exit Sub

    Dim targetCell As Range
    Set targetCell = ws.Range(GetOffsetCellAddress(ws, cellAddress, topRow))

    If targetCell.MergeCells Then
        targetCell.MergeArea.Cells(1, 1).Value = cellValue
    Else
        targetCell.Value = cellValue
    End If
End Sub

Private Function GetOffsetCellAddress(ws As Worksheet, cellAddress As String, topRow As Long) As String
    Dim baseRange As Range
    Set baseRange = ws.Range(cellAddress)
    GetOffsetCellAddress = ws.Cells(baseRange.Row + topRow - 1, baseRange.Column).Address(False, False)
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
        If Len(batchKey) = 0 Then batchKey = CStr(dispatchItem(DispatchItemColumnId))

        If Not groupedBatches.Exists(batchKey) Then groupedBatches.Add batchKey, New Collection

        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)
        batchItems.Add dispatchItem
    Next i

    Set GroupDispatchItemsByBatch = groupedBatches
End Function

Private Function AppendEnvelopeLayoutPage(batchItems As Collection, ByRef firstVisibleSheet As Worksheet, showPreviewGrid As Boolean) As Boolean
    Dim firstItem As Variant
    firstItem = batchItems(1)

    Dim envelopeFormatKey As String
    envelopeFormatKey = LCase$(Trim$(CStr(firstItem(DispatchItemColumnEnvelopeFormatKey))))

    Dim sheetName As String
    sheetName = ResolveEnvelopeLayoutSheetName(envelopeFormatKey)
    If Len(sheetName) = 0 Then Exit Function

    On Error GoTo AppendError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Visible = xlSheetVisible

    Dim topRow As Long
    topRow = GetNextEnvelopeTopRow(ws, envelopeFormatKey)

    If topRow > 1 Then ws.HPageBreaks.Add Before:=ws.Cells(topRow, EnvelopeLayoutFirstColumn)
    EnsureEnvelopeTemplatePageAt ws, envelopeFormatKey, topRow

    RenderEnvelopeLayoutBlock ws, topRow, envelopeFormatKey, batchItems
    If showPreviewGrid Then RenderEnvelopePreviewGrid ws, topRow, envelopeFormatKey
    ConfigureEnvelopePageSettings ws, envelopeFormatKey, topRow + GetEnvelopeRowsPerPage(envelopeFormatKey) - 1

    If firstVisibleSheet Is Nothing Then Set firstVisibleSheet = ws

    AppendEnvelopeLayoutPage = True
    Exit Function

AppendError:
    Debug.Print "AppendEnvelopeLayoutPage error: " & Err.description
    AppendEnvelopeLayoutPage = False
End Function

Private Sub RenderEnvelopePreviewGrid(ws As Worksheet, topRow As Long, envelopeFormatKey As String)
    On Error GoTo GridError

    Dim rowsPerPage As Long
    rowsPerPage = GetEnvelopeRowsPerPage(envelopeFormatKey)

    Dim pageRange As Range
    Set pageRange = ws.Range(ws.Cells(topRow, EnvelopeLayoutFirstColumn), ws.Cells(topRow + rowsPerPage - 1, GetEnvelopeLastColumn(envelopeFormatKey)))
    ApplyPreviewBorder pageRange, RGB(120, 120, 120), xlContinuous, xlThin
    AddPreviewLabel ws, pageRange, t("dispatch.layouts.preview.page", "Envelope page boundary")

    Dim senderRange As Range
    Set senderRange = ws.Range(ws.Cells(topRow + 1, 1), ws.Cells(topRow + 6, 5))
    AddPreviewZone ws, senderRange, t("dispatch.layouts.preview.sender", "Sender zone"), RGB(47, 117, 181)

    Dim outgoingRange As Range
    Set outgoingRange = ws.Range(ws.Cells(topRow + GetOutgoingTopOffset(envelopeFormatKey), 1), ws.Cells(topRow + GetOutgoingBottomOffset(envelopeFormatKey), 5))
    AddPreviewZone ws, outgoingRange, t("dispatch.layouts.preview.outgoing", "Outgoing numbers zone"), RGB(112, 48, 160)

    Dim recipientRange As Range
    Set recipientRange = ws.Range(ws.Cells(topRow + GetRecipientTopOffset(envelopeFormatKey), 6), ws.Cells(topRow + GetPostalCodeTopOffset(envelopeFormatKey) + 1, 11))
    AddPreviewZone ws, recipientRange, t("dispatch.layouts.preview.recipient", "Recipient zone"), RGB(0, 128, 0)

    Dim postalRange As Range
    Set postalRange = ws.Range(ws.Cells(topRow + GetPostalCodeTopOffset(envelopeFormatKey), 8), ws.Cells(topRow + GetPostalCodeTopOffset(envelopeFormatKey) + 1, 11))
    AddPreviewZone ws, postalRange, t("dispatch.layouts.preview.postal_code", "Postal code zone"), RGB(192, 0, 0)

    Exit Sub

GridError:
    Debug.Print "RenderEnvelopePreviewGrid error: " & Err.description
End Sub

Private Sub AddPreviewZone(ws As Worksheet, targetRange As Range, labelText As String, borderColor As Long)
    ApplyPreviewBorder targetRange, borderColor, xlContinuous, xlMedium
    AddPreviewLabel ws, targetRange, labelText
End Sub

Private Sub ApplyPreviewBorder(targetRange As Range, borderColor As Long, lineStyle As Long, lineWeight As Long)
    With targetRange.Borders(xlEdgeLeft)
        .LineStyle = lineStyle
        .Weight = lineWeight
        .Color = borderColor
    End With

    With targetRange.Borders(xlEdgeTop)
        .LineStyle = lineStyle
        .Weight = lineWeight
        .Color = borderColor
    End With

    With targetRange.Borders(xlEdgeRight)
        .LineStyle = lineStyle
        .Weight = lineWeight
        .Color = borderColor
    End With

    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = lineStyle
        .Weight = lineWeight
        .Color = borderColor
    End With
End Sub

Private Sub AddPreviewLabel(ws As Worksheet, targetRange As Range, labelText As String)
    Dim labelShape As Shape
    Set labelShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, targetRange.Left + 2, targetRange.Top + 2, 150, 14)
    labelShape.Name = EnvelopePreviewShapePrefix & CStr(ws.Shapes.count)
    labelShape.TextFrame.Characters.Text = labelText
    labelShape.TextFrame.Characters.Font.Size = 7
    labelShape.TextFrame.Characters.Font.Bold = True
    labelShape.TextFrame.Characters.Font.Color = RGB(80, 80, 80)
    labelShape.Line.Visible = msoFalse
    labelShape.Fill.Visible = msoTrue
    labelShape.Fill.ForeColor.RGB = RGB(255, 255, 204)
    labelShape.Fill.Transparency = 0.2
    labelShape.Placement = xlMoveAndSize
End Sub

Private Function GetNextEnvelopeTopRow(ws As Worksheet, envelopeFormatKey As String) As Long
    Dim rowsPerPage As Long
    rowsPerPage = GetEnvelopeRowsPerPage(envelopeFormatKey)

    Dim markerCell As String
    markerCell = GetBatchMarkerCell(envelopeFormatKey)

    Dim pageStartRow As Long
    pageStartRow = 1

    Do While pageStartRow < ws.Rows.count - rowsPerPage
        If Len(Trim$(CStr(ws.Range(GetOffsetCellAddress(ws, markerCell, pageStartRow)).Value))) = 0 Then
            GetNextEnvelopeTopRow = pageStartRow
            Exit Function
        End If

        pageStartRow = pageStartRow + rowsPerPage
    Loop

    GetNextEnvelopeTopRow = pageStartRow
End Function

Private Sub EnsureEnvelopeTemplatePageAt(ws As Worksheet, envelopeFormatKey As String, topRow As Long)
    Dim rowsPerPage As Long
    rowsPerPage = GetEnvelopeRowsPerPage(envelopeFormatKey)

    If topRow > 1 Then
        Dim templateRange As Range
        Set templateRange = ws.Range(ws.Cells(1, EnvelopeLayoutFirstColumn), ws.Cells(rowsPerPage, GetEnvelopeLastColumn(envelopeFormatKey)))
        templateRange.Copy ws.Cells(topRow, EnvelopeLayoutFirstColumn)

        Dim rowOffset As Long
        For rowOffset = 0 To rowsPerPage - 1
            ws.Rows(topRow + rowOffset).RowHeight = ws.Rows(1 + rowOffset).RowHeight
        Next rowOffset

        ClearEnvelopeRuntimeShapesInPage ws, topRow, rowsPerPage
    End If

    ClearEnvelopeRuntimeShapesInPage ws, topRow, rowsPerPage
    ClearEnvelopeTemplateDynamicCells ws, envelopeFormatKey, topRow
End Sub

Private Sub ConfigureEnvelopeLayoutGrid(ws As Worksheet, envelopeFormatKey As String)
    Dim colIndex As Long
    For colIndex = EnvelopeLayoutFirstColumn To GetEnvelopeLastColumn(envelopeFormatKey)
        ws.Columns(colIndex).ColumnWidth = 12
    Next colIndex

    Select Case envelopeFormatKey
    Case "c4"
        ws.Columns("A").ColumnWidth = 10
        ws.Columns("B").ColumnWidth = 6
        ws.Columns("C").ColumnWidth = 16
        ws.Columns("D").ColumnWidth = 13
        ws.Columns("E").ColumnWidth = 11
        ws.Columns("F").ColumnWidth = 9
        ws.Columns("G").ColumnWidth = 10
        ws.Columns("H").ColumnWidth = 18
        ws.Columns("I").ColumnWidth = 18
        ws.Columns("J").ColumnWidth = 18
        ws.Columns("K").ColumnWidth = 12
    Case "c5"
        ws.Columns("A").ColumnWidth = 10
        ws.Columns("B").ColumnWidth = 13
        ws.Columns("C").ColumnWidth = 13
        ws.Columns("D").ColumnWidth = 16
        ws.Columns("E").ColumnWidth = 16
        ws.Columns("F").ColumnWidth = 16
    Case Else
        ws.Columns("A").ColumnWidth = 9
        ws.Columns("B").ColumnWidth = 11
        ws.Columns("C").ColumnWidth = 11
        ws.Columns("D").ColumnWidth = 14
        ws.Columns("E").ColumnWidth = 14
        ws.Columns("F").ColumnWidth = 14
    End Select
End Sub

Private Sub RenderEnvelopeLayoutBlock(ws As Worksheet, topRow As Long, envelopeFormatKey As String, batchItems As Collection)
    RenderEnvelopeTemplateBlock ws, topRow, envelopeFormatKey, batchItems
    Exit Sub

    If envelopeFormatKey = "c4" Then
        RenderEnvelopeLayoutBlockC4 ws, topRow, batchItems
        Exit Sub
    End If

    Dim firstItem As Variant
    firstItem = batchItems(1)

    Dim senderName As String
    senderName = Trim$(CStr(firstItem(DispatchItemColumnSenderName)))

    Dim senderBlock As String
    senderBlock = BuildEnvelopeSenderBlock(senderName)

    Dim outgoingText As String
    outgoingText = BuildBatchOutgoingNumbersText(batchItems)

    Dim recipientBlock As String
    recipientBlock = BuildEnvelopeRecipientBlock(CStr(firstItem(DispatchItemColumnAddressee)), CStr(firstItem(DispatchItemColumnAddressLine)), CStr(firstItem(DispatchItemColumnPostalCode)))

    Dim rowsPerPage As Long
    rowsPerPage = GetEnvelopeRowsPerPage(envelopeFormatKey)

    Dim blockRange As Range
    Set blockRange = ws.Range(ws.Cells(topRow, EnvelopeLayoutFirstColumn), ws.Cells(topRow + rowsPerPage - 1, GetEnvelopeLastColumn(envelopeFormatKey)))
    blockRange.Clear
    blockRange.Font.Name = "Times New Roman"
    blockRange.Font.Size = GetEnvelopeBaseFontSize(envelopeFormatKey)
    blockRange.VerticalAlignment = xlTop
    blockRange.WrapText = True

    Dim senderRange As Range
    Set senderRange = ws.Range(ws.Cells(topRow + 1, 1), ws.Cells(topRow + 3, 3))
    senderRange.Merge
    senderRange.Value = senderBlock
    senderRange.Font.Size = GetEnvelopeSmallFontSize(envelopeFormatKey)
    senderRange.HorizontalAlignment = xlLeft

    Dim outgoingRange As Range
    Set outgoingRange = ws.Range(ws.Cells(topRow + 4, 1), ws.Cells(topRow + 6, 3))
    outgoingRange.Merge
    outgoingRange.Value = outgoingText
    outgoingRange.Font.Size = GetEnvelopeSmallFontSize(envelopeFormatKey)
    outgoingRange.HorizontalAlignment = xlLeft

    Dim recipientRange As Range
    Set recipientRange = ws.Range(ws.Cells(topRow + GetRecipientTopOffset(envelopeFormatKey), 4), ws.Cells(topRow + GetRecipientTopOffset(envelopeFormatKey) + GetRecipientBlockHeight(envelopeFormatKey), 6))
    recipientRange.Merge
    recipientRange.Value = recipientBlock
    recipientRange.Font.Size = GetEnvelopeBaseFontSize(envelopeFormatKey)
    recipientRange.HorizontalAlignment = xlLeft
    recipientRange.VerticalAlignment = xlTop

    Dim postalRange As Range
    Set postalRange = ws.Range(ws.Cells(topRow + GetPostalCodeTopOffset(envelopeFormatKey), 4), ws.Cells(topRow + GetPostalCodeTopOffset(envelopeFormatKey), 6))
    postalRange.Merge
    postalRange.Value = CStr(firstItem(DispatchItemColumnPostalCode))
    postalRange.Font.Size = GetEnvelopePostalFontSize(envelopeFormatKey)
    postalRange.Font.Bold = True
    postalRange.HorizontalAlignment = xlLeft

    Dim batchRange As Range
    Set batchRange = ws.Range(ws.Cells(topRow + rowsPerPage - 1, 1), ws.Cells(topRow + rowsPerPage - 1, 3))
    batchRange.Merge
    batchRange.Value = BuildEnvelopeBatchMarker(firstItem)
    batchRange.Font.Size = 7
    batchRange.Font.Color = RGB(255, 255, 255)

    Dim rowRangeAddress As String
    rowRangeAddress = CStr(topRow) & ":" & CStr(topRow + rowsPerPage - 1)
    ws.Rows(rowRangeAddress).RowHeight = GetEnvelopeRowHeight(envelopeFormatKey)
End Sub

Private Sub RenderEnvelopeTemplateBlock(ws As Worksheet, topRow As Long, envelopeFormatKey As String, batchItems As Collection)
    Dim firstItem As Variant
    firstItem = batchItems(1)

    ClearEnvelopeTemplateDynamicCells ws, envelopeFormatKey, topRow

    Dim senderName As String
    senderName = Trim$(CStr(firstItem(DispatchItemColumnSenderName)))

    Dim senderPostalCode As String
    senderPostalCode = DispatchRepositoryGetSenderPostalCode(senderName)

    Dim recipientPostalCode As String
    recipientPostalCode = Trim$(CStr(firstItem(DispatchItemColumnPostalCode)))

    Dim senderLines As Variant
    senderLines = BuildEnvelopeAddressLines(DispatchRepositoryGetSenderAddressBlock(senderName), senderPostalCode, 3)

    Dim recipientLines As Variant
    recipientLines = BuildEnvelopeAddressLines(CStr(firstItem(DispatchItemColumnAddressLine)), recipientPostalCode, 3)

    SetTemplateValue ws, topRow, GetSenderNameCell(envelopeFormatKey), senderName
    SetTemplateValue ws, topRow, GetSenderAddressCell(envelopeFormatKey, 1), CStr(senderLines(1))
    SetTemplateValue ws, topRow, GetSenderAddressCell(envelopeFormatKey, 2), CStr(senderLines(2))
    SetTemplateValue ws, topRow, GetSenderAddressCell(envelopeFormatKey, 3), CStr(senderLines(3))
    SetTemplateValue ws, topRow, GetSenderPostalCell(envelopeFormatKey), senderPostalCode
    SetEnvelopeOutgoingNumbers ws, topRow, envelopeFormatKey, BuildBatchOutgoingNumbersText(batchItems)
    SetTemplateValue ws, topRow, GetRecipientNameCell(envelopeFormatKey), Trim$(CStr(firstItem(DispatchItemColumnAddressee)))
    SetTemplateValue ws, topRow, GetRecipientAddressCell(envelopeFormatKey, 1), CStr(recipientLines(1))
    SetTemplateValue ws, topRow, GetRecipientAddressCell(envelopeFormatKey, 2), CStr(recipientLines(2))
    SetTemplateValue ws, topRow, GetRecipientAddressCell(envelopeFormatKey, 3), CStr(recipientLines(3))
    SetTemplateValue ws, topRow, GetRecipientPostalCell(envelopeFormatKey), recipientPostalCode
    SetEnvelopeMailTypeMark ws, topRow, envelopeFormatKey, CStr(firstItem(DispatchItemColumnMailType))
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetLargePostalCell(envelopeFormatKey)
    RenderLeftPostalIndexGuide ws, topRow, envelopeFormatKey, recipientPostalCode
    SetTemplateValue ws, topRow, GetBatchMarkerCell(envelopeFormatKey), BuildEnvelopeBatchMarker(firstItem)
    ws.Range(GetOffsetCellAddress(ws, GetBatchMarkerCell(envelopeFormatKey), topRow)).Font.Color = RGB(255, 255, 255)
End Sub

Private Sub SetEnvelopeOutgoingNumbers(ws As Worksheet, topRow As Long, envelopeFormatKey As String, outgoingText As String)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetOutgoingCell(envelopeFormatKey)

    Dim anchorRange As Range
    Set anchorRange = ws.Range(GetOffsetCellAddress(ws, GetOutgoingCell(envelopeFormatKey), topRow)).MergeArea

    Dim shapeWidth As Double
    shapeWidth = 260

    Dim shapeHeight As Double
    shapeHeight = 64

    Dim shapeLeftOffset As Double
    shapeLeftOffset = 6

    Dim shapeTopOffset As Double
    shapeTopOffset = 4

    Dim fontSize As Double
    fontSize = 10

    Select Case envelopeFormatKey
    Case "c4"
        shapeWidth = 360
        shapeHeight = 82
        shapeTopOffset = 2
        fontSize = 11
    Case "dl"
        shapeWidth = 240
        shapeHeight = 52
        shapeTopOffset = 2
        fontSize = 9
    End Select

    Dim outgoingShape As Shape
    Set outgoingShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, anchorRange.Left + shapeLeftOffset, anchorRange.Top + shapeTopOffset, shapeWidth, shapeHeight)
    outgoingShape.Name = EnvelopeDynamicShapePrefix & "Outgoing_" & envelopeFormatKey & "_" & CStr(topRow)
    outgoingShape.TextFrame.Characters.Text = outgoingText
    outgoingShape.TextFrame.Characters.Font.Name = "Times New Roman"
    outgoingShape.TextFrame.Characters.Font.Size = fontSize
    outgoingShape.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    outgoingShape.TextFrame.HorizontalAlignment = xlHAlignLeft
    outgoingShape.TextFrame.VerticalAlignment = xlVAlignTop
    outgoingShape.Line.Visible = msoFalse
    outgoingShape.Fill.Visible = msoFalse
    outgoingShape.Placement = xlMoveAndSize
End Sub

Private Sub SetEnvelopeMailTypeMark(ws As Worksheet, topRow As Long, envelopeFormatKey As String, mailType As String)
    Dim markText As String
    markText = BuildEnvelopeMailTypeMark(mailType)

    ClearTemplateCell ws, envelopeFormatKey, topRow, GetMailTypeMarkCell(envelopeFormatKey, 1)
    ClearTemplateCell ws, envelopeFormatKey, topRow, GetMailTypeMarkCell(envelopeFormatKey, 2)
    If Len(markText) = 0 Then Exit Sub

    Dim markLines As Variant
    markLines = Split(markText, vbCrLf)

    SetTemplateValue ws, topRow, GetMailTypeMarkCell(envelopeFormatKey, 1), CStr(markLines(0))
    FormatMailTypeMarkCell ws, topRow, GetMailTypeMarkCell(envelopeFormatKey, 1)

    If UBound(markLines) >= 1 Then
        SetTemplateValue ws, topRow, GetMailTypeMarkCell(envelopeFormatKey, 2), CStr(markLines(1))
        FormatMailTypeMarkCell ws, topRow, GetMailTypeMarkCell(envelopeFormatKey, 2)
    End If
End Sub

Private Sub FormatMailTypeMarkCell(ws As Worksheet, topRow As Long, cellAddress As String)
    If Len(cellAddress) = 0 Then Exit Sub

    Dim targetCell As Range
    Set targetCell = ws.Range(GetOffsetCellAddress(ws, cellAddress, topRow))
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)

    targetCell.Font.Bold = True
    targetCell.Font.Size = 11
    targetCell.HorizontalAlignment = xlCenter
    targetCell.VerticalAlignment = xlCenter
End Sub

Private Function BuildEnvelopeMailTypeMark(mailType As String) As String
    Dim normalizedMailType As String
    normalizedMailType = UCase$(Trim$(mailType))

    If Len(normalizedMailType) = 0 Then
        BuildEnvelopeMailTypeMark = GetMailTypeRegisteredMark()
        Exit Function
    End If

    If normalizedMailType = "SIMPLE" Or normalizedMailType = "ORDINARY" Or normalizedMailType = GetMailTypeSimpleWord() Or normalizedMailType = GetMailTypeOrdinaryWord() Then Exit Function

    If normalizedMailType = "REGISTERED_NOTICE" Or normalizedMailType = "REGISTERED_WITH_NOTICE" Then
        BuildEnvelopeMailTypeMark = GetMailTypeRegisteredMark() & vbCrLf & GetMailTypeNoticeMark()
        Exit Function
    End If

    If normalizedMailType = "DECLARED_VALUE_NOTICE" Or normalizedMailType = "VALUE_NOTICE" Then
        BuildEnvelopeMailTypeMark = GetMailTypeDeclaredValueMark() & vbCrLf & GetMailTypeNoticeMark()
        Exit Function
    End If

    If normalizedMailType = "DECLARED_VALUE" Or normalizedMailType = "VALUE" Then
        BuildEnvelopeMailTypeMark = GetMailTypeDeclaredValueMark()
        Exit Function
    End If

    If InStr(1, normalizedMailType, GetMailTypeValueStem(), vbTextCompare) > 0 Then
        BuildEnvelopeMailTypeMark = GetMailTypeDeclaredValueMark()
        If InStr(1, normalizedMailType, GetMailTypeNoticeStem(), vbTextCompare) > 0 Then BuildEnvelopeMailTypeMark = BuildEnvelopeMailTypeMark & vbCrLf & GetMailTypeNoticeMark()
        Exit Function
    End If

    If InStr(1, normalizedMailType, GetMailTypeNoticeStem(), vbTextCompare) > 0 Then
        BuildEnvelopeMailTypeMark = GetMailTypeRegisteredMark() & vbCrLf & GetMailTypeNoticeMark()
        Exit Function
    End If

    If normalizedMailType = "REGISTERED" Or normalizedMailType = GetMailTypeRegisteredMark() Then
        BuildEnvelopeMailTypeMark = GetMailTypeRegisteredMark()
        Exit Function
    End If

    BuildEnvelopeMailTypeMark = normalizedMailType
End Function

Private Function GetMailTypeRegisteredMark() As String
    GetMailTypeRegisteredMark = ChrW$(1047) & ChrW$(1040) & ChrW$(1050) & ChrW$(1040) & ChrW$(1047) & ChrW$(1053) & ChrW$(1054) & ChrW$(1045)
End Function

Private Function GetMailTypeNoticeMark() As String
    GetMailTypeNoticeMark = ChrW$(1057) & " " & ChrW$(1059) & ChrW$(1042) & ChrW$(1045) & ChrW$(1044) & ChrW$(1054) & ChrW$(1052) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053) & ChrW$(1048) & ChrW$(1045) & ChrW$(1052)
End Function

Private Function GetMailTypeDeclaredValueMark() As String
    GetMailTypeDeclaredValueMark = ChrW$(1057) & " " & ChrW$(1054) & ChrW$(1041) & ChrW$(1066) & ChrW$(1071) & ChrW$(1042) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053) & ChrW$(1053) & ChrW$(1054) & ChrW$(1049) & " " & ChrW$(1062) & ChrW$(1045) & ChrW$(1053) & ChrW$(1053) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058) & ChrW$(1068) & ChrW$(1070)
End Function

Private Function GetMailTypeSimpleWord() As String
    GetMailTypeSimpleWord = ChrW$(1055) & ChrW$(1056) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058) & ChrW$(1054) & ChrW$(1045)
End Function

Private Function GetMailTypeOrdinaryWord() As String
    GetMailTypeOrdinaryWord = ChrW$(1054) & ChrW$(1041) & ChrW$(1067) & ChrW$(1063) & ChrW$(1053) & ChrW$(1054) & ChrW$(1045)
End Function

Private Function GetMailTypeNoticeStem() As String
    GetMailTypeNoticeStem = ChrW$(1059) & ChrW$(1042) & ChrW$(1045) & ChrW$(1044) & ChrW$(1054) & ChrW$(1052) & ChrW$(1051) & ChrW$(1045) & ChrW$(1053)
End Function

Private Function GetMailTypeValueStem() As String
    GetMailTypeValueStem = ChrW$(1062) & ChrW$(1045) & ChrW$(1053) & ChrW$(1053) & ChrW$(1054) & ChrW$(1057) & ChrW$(1058)
End Function

Private Sub RenderLeftPostalIndexGuide(ws As Worksheet, topRow As Long, envelopeFormatKey As String, postalCode As String)
    On Error GoTo GuideError

    Dim normalizedPostalCode As String
    normalizedPostalCode = OnlyDigits(postalCode)
    If Len(normalizedPostalCode) = 0 Then Exit Sub

    Dim targetRange As Range
    Set targetRange = ws.Range(GetOffsetCellAddress(ws, GetLargePostalCell(envelopeFormatKey), topRow)).MergeArea

    Dim startLeft As Double
    startLeft = targetRange.Left + 9

    Dim startTop As Double
    startTop = targetRange.Top + 9

    Dim digitWidth As Double
    digitWidth = 23

    Dim digitStep As Double
    digitStep = 22

    Dim digitHeight As Double
    digitHeight = 29

    Dim strokeWidth As Double
    strokeWidth = 2.2

    Select Case envelopeFormatKey
    Case "c4"
        startLeft = targetRange.Left + 11
        startTop = targetRange.Top - 24
        digitWidth = 28
        digitStep = 26
        digitHeight = 34
        strokeWidth = 2.5
    Case "dl"
        startLeft = targetRange.Left + 7
        startTop = targetRange.Top + 4
        digitWidth = 18
        digitStep = 18
        digitHeight = 24
        strokeWidth = 1.8
    End Select

    Dim digitIndex As Long
    AddPostalGuideStartMarker ws, envelopeFormatKey, startLeft - 1, startTop + 7, topRow

    For digitIndex = 1 To Len(normalizedPostalCode)
        Dim digitLeft As Double
        digitLeft = startLeft + (digitIndex * digitStep)

        AddPostalGuideDigit ws, envelopeFormatKey, Mid$(normalizedPostalCode, digitIndex, 1), digitLeft, startTop + 8, digitWidth, digitHeight, strokeWidth, digitIndex, topRow
    Next digitIndex

    For digitIndex = 1 To 7
        AddPostalGuideBar ws, envelopeFormatKey, startLeft + ((digitIndex - 1) * digitStep), startTop, digitIndex, topRow
    Next digitIndex

    Exit Sub

GuideError:
    Debug.Print "RenderLeftPostalIndexGuide error: " & Err.description
End Sub

Private Sub AddPostalGuideStartMarker(ws As Worksheet, envelopeFormatKey As String, leftPosition As Double, topPosition As Double, topRow As Long)
    Dim markerShape As Shape
    Set markerShape = ws.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition, 18, 3)
    markerShape.Name = EnvelopeDynamicShapePrefix & "PostalStart_" & envelopeFormatKey & "_" & CStr(topRow)
    markerShape.Line.Visible = msoFalse
    markerShape.Fill.Visible = msoTrue
    markerShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
    markerShape.Placement = xlMoveAndSize
End Sub

Private Sub AddPostalGuideBar(ws As Worksheet, envelopeFormatKey As String, leftPosition As Double, topPosition As Double, digitIndex As Long, topRow As Long)
    Dim barShape As Shape
    Set barShape = ws.Shapes.AddShape(msoShapeRectangle, leftPosition + 1, topPosition, 18, 4)
    barShape.Name = EnvelopeDynamicShapePrefix & "PostalBar_" & envelopeFormatKey & "_" & CStr(topRow) & "_" & CStr(digitIndex)
    barShape.Line.Visible = msoFalse
    barShape.Fill.Visible = msoTrue
    barShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
    barShape.Placement = xlMoveAndSize
End Sub

Private Sub AddPostalGuideDigit(ws As Worksheet, envelopeFormatKey As String, digitText As String, leftPosition As Double, topPosition As Double, digitWidth As Double, digitHeight As Double, strokeWidth As Double, digitIndex As Long, topRow As Long)
    Dim normalizedDigit As String
    normalizedDigit = Left$(OnlyDigits(digitText), 1)
    If Len(normalizedDigit) = 0 Then Exit Sub

    Dim segmentLeft As Double
    segmentLeft = leftPosition + 4

    Dim segmentTop As Double
    segmentTop = topPosition + 3

    Dim segmentWidth As Double
    segmentWidth = digitWidth - 10

    AddPostalDigitSegments ws, envelopeFormatKey, normalizedDigit, segmentLeft, segmentTop, segmentWidth, digitHeight, strokeWidth, digitIndex, topRow
End Sub

Private Sub AddPostalDigitSegments(ws As Worksheet, envelopeFormatKey As String, digitText As String, leftPosition As Double, topPosition As Double, digitWidth As Double, digitHeight As Double, strokeWidth As Double, digitIndex As Long, topRow As Long)
    If digitText = "1" Then
        AddPostalDigitOneSegments ws, envelopeFormatKey, leftPosition, topPosition, digitWidth, digitHeight, strokeWidth, digitIndex, topRow
        Exit Sub
    End If

    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "A"), leftPosition + strokeWidth, topPosition, digitWidth - (strokeWidth * 2), strokeWidth, "A", digitIndex, topRow
    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "B"), leftPosition + digitWidth - strokeWidth, topPosition + strokeWidth, strokeWidth, (digitHeight / 2) - strokeWidth, "B", digitIndex, topRow
    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "C"), leftPosition + digitWidth - strokeWidth, topPosition + (digitHeight / 2), strokeWidth, (digitHeight / 2) - strokeWidth, "C", digitIndex, topRow
    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "D"), leftPosition + strokeWidth, topPosition + digitHeight - strokeWidth, digitWidth - (strokeWidth * 2), strokeWidth, "D", digitIndex, topRow
    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "E"), leftPosition, topPosition + (digitHeight / 2), strokeWidth, (digitHeight / 2) - strokeWidth, "E", digitIndex, topRow
    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "F"), leftPosition, topPosition + strokeWidth, strokeWidth, (digitHeight / 2) - strokeWidth, "F", digitIndex, topRow
    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, ShouldPostalDigitDrawSegment(digitText, "G"), leftPosition + strokeWidth, topPosition + (digitHeight / 2) - (strokeWidth / 2), digitWidth - (strokeWidth * 2), strokeWidth, "G", digitIndex, topRow
End Sub

Private Sub AddPostalDigitOneSegments(ws As Worksheet, envelopeFormatKey As String, leftPosition As Double, topPosition As Double, digitWidth As Double, digitHeight As Double, strokeWidth As Double, digitIndex As Long, topRow As Long)
    Dim verticalLeft As Double
    verticalLeft = leftPosition + digitWidth - strokeWidth

    Dim verticalTop As Double
    verticalTop = topPosition + strokeWidth

    Dim verticalHeight As Double
    verticalHeight = digitHeight - strokeWidth

    AddPostalDigitSegmentIfNeeded ws, envelopeFormatKey, True, verticalLeft, verticalTop, strokeWidth, verticalHeight, "C", digitIndex, topRow
    AddPostalDigitDiagonalSegment ws, envelopeFormatKey, leftPosition + (digitWidth * 0.45), topPosition + (digitHeight * 0.28), verticalLeft + (strokeWidth / 2), verticalTop, strokeWidth, digitIndex, topRow
End Sub

Private Sub AddPostalDigitDiagonalSegment(ws As Worksheet, envelopeFormatKey As String, startLeft As Double, startTop As Double, endLeft As Double, endTop As Double, strokeWidth As Double, digitIndex As Long, topRow As Long)
    Dim segmentShape As Shape
    Set segmentShape = ws.Shapes.AddLine(startLeft, startTop, endLeft, endTop)
    segmentShape.Name = EnvelopeDynamicShapePrefix & "PostalDigit_" & envelopeFormatKey & "_" & CStr(topRow) & "_" & CStr(digitIndex) & "_SLANT"
    segmentShape.Line.Visible = msoTrue
    segmentShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    segmentShape.Line.Weight = strokeWidth
    segmentShape.Placement = xlMoveAndSize
End Sub

Private Sub AddPostalDigitSegmentIfNeeded(ws As Worksheet, envelopeFormatKey As String, shouldDraw As Boolean, leftPosition As Double, topPosition As Double, segmentWidth As Double, segmentHeight As Double, segmentCode As String, digitIndex As Long, topRow As Long)
    If Not shouldDraw Then Exit Sub

    Dim segmentShape As Shape
    Set segmentShape = ws.Shapes.AddShape(msoShapeRectangle, leftPosition, topPosition, segmentWidth, segmentHeight)
    segmentShape.Name = EnvelopeDynamicShapePrefix & "PostalDigit_" & envelopeFormatKey & "_" & CStr(topRow) & "_" & CStr(digitIndex) & "_" & segmentCode
    segmentShape.Line.Visible = msoFalse
    segmentShape.Fill.Visible = msoTrue
    segmentShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
    segmentShape.Placement = xlMoveAndSize
End Sub

Private Function ShouldPostalDigitDrawSegment(digitText As String, segmentCode As String) As Boolean
    Select Case digitText
    Case "0"
        ShouldPostalDigitDrawSegment = InStr(1, "ABCDEF", segmentCode, vbTextCompare) > 0
    Case "1"
        ShouldPostalDigitDrawSegment = InStr(1, "BC", segmentCode, vbTextCompare) > 0
    Case "2"
        ShouldPostalDigitDrawSegment = InStr(1, "ABDEG", segmentCode, vbTextCompare) > 0
    Case "3"
        ShouldPostalDigitDrawSegment = InStr(1, "ABCDG", segmentCode, vbTextCompare) > 0
    Case "4"
        ShouldPostalDigitDrawSegment = InStr(1, "BCFG", segmentCode, vbTextCompare) > 0
    Case "5"
        ShouldPostalDigitDrawSegment = InStr(1, "ACDFG", segmentCode, vbTextCompare) > 0
    Case "6"
        ShouldPostalDigitDrawSegment = InStr(1, "ACDEFG", segmentCode, vbTextCompare) > 0
    Case "7"
        ShouldPostalDigitDrawSegment = InStr(1, "ABC", segmentCode, vbTextCompare) > 0
    Case "8"
        ShouldPostalDigitDrawSegment = InStr(1, "ABCDEFG", segmentCode, vbTextCompare) > 0
    Case "9"
        ShouldPostalDigitDrawSegment = InStr(1, "ABCDFG", segmentCode, vbTextCompare) > 0
    End Select
End Function

Private Sub RenderEnvelopeLayoutBlockC4(ws As Worksheet, topRow As Long, batchItems As Collection)
    Dim firstItem As Variant
    firstItem = batchItems(1)

    Dim senderName As String
    senderName = Trim$(CStr(firstItem(DispatchItemColumnSenderName)))

    Dim senderAddress As String
    senderAddress = DispatchRepositoryGetSenderAddressBlock(senderName)

    Dim senderPostalCode As String
    senderPostalCode = DispatchRepositoryGetSenderPostalCode(senderName)

    Dim recipientLine1 As String
    Dim recipientLine2 As String
    SplitEnvelopeAddressLine CStr(firstItem(DispatchItemColumnAddressLine)), CStr(firstItem(DispatchItemColumnPostalCode)), recipientLine1, recipientLine2

    Dim rowsPerPage As Long
    rowsPerPage = GetEnvelopeRowsPerPage("c4")

    Dim blockRange As Range
    Set blockRange = ws.Range(ws.Cells(topRow, EnvelopeLayoutFirstColumn), ws.Cells(topRow + rowsPerPage - 1, EnvelopeLayoutLastColumnC4))
    blockRange.Clear
    blockRange.Font.Name = "Times New Roman"
    blockRange.Font.Size = GetEnvelopeBaseFontSize("c4")
    blockRange.VerticalAlignment = xlTop
    blockRange.WrapText = True

    ws.Range(ws.Cells(topRow + 1, 1), ws.Cells(topRow + 1, 2)).Merge
    ws.Cells(topRow + 1, 1).Value = t("dispatch.layouts.print.sender_name_label", "From")
    ws.Cells(topRow + 1, 1).Font.Bold = True
    ws.Range(ws.Cells(topRow + 1, 3), ws.Cells(topRow + 1, 11)).Merge
    ws.Cells(topRow + 1, 3).Value = senderName
    ws.Cells(topRow + 1, 3).Font.Italic = True

    ws.Range(ws.Cells(topRow + 2, 1), ws.Cells(topRow + 2, 2)).Merge
    ws.Cells(topRow + 2, 1).Value = t("dispatch.layouts.print.sender_address_label", "Sender address")
    ws.Cells(topRow + 2, 1).Font.Bold = True
    ws.Range(ws.Cells(topRow + 2, 3), ws.Cells(topRow + 4, 11)).Merge
    ws.Cells(topRow + 2, 3).Value = senderAddress
    ws.Cells(topRow + 2, 3).Font.Italic = True

    ws.Range(ws.Cells(topRow + 5, 5), ws.Cells(topRow + 5, 7)).Merge
    ws.Cells(topRow + 5, 5).Value = t("dispatch.layouts.print.sender_index_label", "Sender postal code")
    ws.Cells(topRow + 5, 5).Font.Size = 7
    ws.Cells(topRow + 5, 5).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(topRow + 6, 5), ws.Cells(topRow + 6, 7)).Merge
    ws.Cells(topRow + 6, 5).Value = senderPostalCode
    ws.Cells(topRow + 6, 5).Font.Size = 14
    ws.Cells(topRow + 6, 5).Font.Bold = True
    ws.Cells(topRow + 6, 5).HorizontalAlignment = xlCenter

    Dim outgoingRange As Range
    Set outgoingRange = ws.Range(ws.Cells(topRow + 8, 1), ws.Cells(topRow + 12, 5))
    outgoingRange.Merge
    outgoingRange.Value = BuildBatchOutgoingNumbersText(batchItems)
    outgoingRange.Font.Size = GetEnvelopeSmallFontSize("c4")
    outgoingRange.HorizontalAlignment = xlLeft

    ws.Range(ws.Cells(topRow + 17, 6), ws.Cells(topRow + 17, 7)).Merge
    ws.Cells(topRow + 17, 6).Value = t("dispatch.layouts.print.recipient_name_label", "To")
    ws.Cells(topRow + 17, 6).Font.Bold = True
    ws.Range(ws.Cells(topRow + 17, 8), ws.Cells(topRow + 17, 11)).Merge
    ws.Cells(topRow + 17, 8).Value = Trim$(CStr(firstItem(DispatchItemColumnAddressee)))
    ws.Cells(topRow + 17, 8).Font.Italic = True

    ws.Range(ws.Cells(topRow + 19, 6), ws.Cells(topRow + 19, 7)).Merge
    ws.Cells(topRow + 19, 6).Value = t("dispatch.layouts.print.recipient_address_label", "Recipient address")
    ws.Cells(topRow + 19, 6).Font.Bold = True
    ws.Range(ws.Cells(topRow + 19, 8), ws.Cells(topRow + 19, 11)).Merge
    ws.Cells(topRow + 19, 8).Value = recipientLine1
    ws.Cells(topRow + 19, 8).Font.Italic = True
    ws.Range(ws.Cells(topRow + 20, 8), ws.Cells(topRow + 20, 11)).Merge
    ws.Cells(topRow + 20, 8).Value = recipientLine2
    ws.Cells(topRow + 20, 8).Font.Italic = True

    ws.Range(ws.Cells(topRow + 22, 1), ws.Cells(topRow + 23, 3)).Merge
    ws.Cells(topRow + 22, 1).Value = "-" & Trim$(CStr(firstItem(DispatchItemColumnPostalCode)))
    ws.Cells(topRow + 22, 1).Font.Size = 28
    ws.Cells(topRow + 22, 1).HorizontalAlignment = xlCenter
    ws.Cells(topRow + 22, 1).VerticalAlignment = xlCenter

    ws.Range(ws.Cells(topRow + 22, 8), ws.Cells(topRow + 22, 11)).Merge
    ws.Cells(topRow + 22, 8).Value = t("dispatch.layouts.print.recipient_index_label", "Recipient postal code")
    ws.Cells(topRow + 22, 8).Font.Size = 7
    ws.Cells(topRow + 22, 8).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(topRow + 23, 8), ws.Cells(topRow + 23, 11)).Merge
    ws.Cells(topRow + 23, 8).Value = Trim$(CStr(firstItem(DispatchItemColumnPostalCode)))
    ws.Cells(topRow + 23, 8).Font.Size = GetEnvelopePostalFontSize("c4")
    ws.Cells(topRow + 23, 8).Font.Bold = True
    ws.Cells(topRow + 23, 8).HorizontalAlignment = xlCenter

    Dim batchRange As Range
    Set batchRange = ws.Range(ws.Cells(topRow + rowsPerPage - 1, 1), ws.Cells(topRow + rowsPerPage - 1, 3))
    batchRange.Merge
    batchRange.Value = BuildEnvelopeBatchMarker(firstItem)
    batchRange.Font.Size = 7
    batchRange.Font.Color = RGB(255, 255, 255)

    Dim rowRangeAddress As String
    rowRangeAddress = CStr(topRow) & ":" & CStr(topRow + rowsPerPage - 1)
    ws.Rows(rowRangeAddress).RowHeight = GetEnvelopeRowHeight("c4")
End Sub

Private Sub ConfigureEnvelopePageSettings(ws As Worksheet, envelopeFormatKey As String, printedLastRow As Long)
    On Error GoTo PageSetupError

    Dim lastRow As Long
    lastRow = printedLastRow
    If lastRow < 1 Then lastRow = 1

    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, EnvelopeLayoutFirstColumn), ws.Cells(lastRow, GetEnvelopeLastColumn(envelopeFormatKey))).Address

    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = GetEnvelopePaperSize(envelopeFormatKey)
        If GetEnvelopeZoomPercent(envelopeFormatKey) > 0 Then
            .Zoom = GetEnvelopeZoomPercent(envelopeFormatKey)
        Else
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End If
        .LeftMargin = Application.CentimetersToPoints(GetEnvelopeLeftMarginCm(envelopeFormatKey))
        .RightMargin = Application.CentimetersToPoints(GetEnvelopeRightMarginCm(envelopeFormatKey))
        .TopMargin = Application.CentimetersToPoints(GetEnvelopeTopMarginCm(envelopeFormatKey))
        .BottomMargin = Application.CentimetersToPoints(GetEnvelopeBottomMarginCm(envelopeFormatKey))
        .CenterHorizontally = True
        .CenterVertically = True
    End With

    Exit Sub

PageSetupError:
    Debug.Print "ConfigureEnvelopePageSettings error: " & Err.description
End Sub

Private Sub FinalizeEnvelopeLayoutSheet(sheetName As String)
    On Error GoTo FinalizeError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim envelopeFormatKey As String
    envelopeFormatKey = ResolveEnvelopeFormatKeyFromSheetName(sheetName)

    If Len(envelopeFormatKey) > 0 Then
        If Len(Trim$(CStr(ws.Range(GetBatchMarkerCell(envelopeFormatKey)).Value))) = 0 Then
            ws.Visible = xlSheetVeryHidden
            Exit Sub
        End If
    End If

    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        ws.Visible = xlSheetVeryHidden
        Exit Sub
    End If

    ws.Visible = xlSheetVisible
    Exit Sub

FinalizeError:
    Debug.Print "FinalizeEnvelopeLayoutSheet error: " & Err.description
End Sub

Private Function BuildEnvelopeSenderBlock(senderName As String) As String
    BuildEnvelopeSenderBlock = senderName

    Dim senderAddress As String
    senderAddress = DispatchRepositoryGetSenderAddressBlock(senderName)

    If Len(Trim$(senderAddress)) > 0 Then
        If Len(Trim$(BuildEnvelopeSenderBlock)) > 0 Then BuildEnvelopeSenderBlock = BuildEnvelopeSenderBlock & vbCrLf
        BuildEnvelopeSenderBlock = BuildEnvelopeSenderBlock & senderAddress
    End If
End Function

Private Function BuildEnvelopeRecipientBlock(addresseeText As String, addressLine As String, postalCode As String) As String
    BuildEnvelopeRecipientBlock = Trim$(addresseeText)

    If Len(Trim$(addressLine)) > 0 Then
        If Len(BuildEnvelopeRecipientBlock) > 0 Then BuildEnvelopeRecipientBlock = BuildEnvelopeRecipientBlock & vbCrLf
        BuildEnvelopeRecipientBlock = BuildEnvelopeRecipientBlock & Trim$(addressLine)
    End If

    If Len(Trim$(postalCode)) > 0 Then
        If InStr(1, addressLine, Trim$(postalCode), vbTextCompare) = 0 Then
            BuildEnvelopeRecipientBlock = BuildEnvelopeRecipientBlock & vbCrLf & Trim$(postalCode)
        End If
    End If
End Function

Private Sub SplitEnvelopeAddressLine(addressLine As String, postalCode As String, ByRef firstLine As String, ByRef secondLine As String)
    Dim normalizedAddress As String
    normalizedAddress = Trim$(addressLine)

    If Len(Trim$(postalCode)) > 0 Then
        normalizedAddress = Replace(normalizedAddress, ", " & Trim$(postalCode), "", 1, -1, vbTextCompare)
        normalizedAddress = Replace(normalizedAddress, Trim$(postalCode), "", 1, -1, vbTextCompare)
    End If

    normalizedAddress = Trim$(normalizedAddress)
    If Len(normalizedAddress) <= 48 Then
        firstLine = normalizedAddress
        secondLine = ""
        Exit Sub
    End If

    Dim splitPosition As Long
    splitPosition = InStrRev(Left$(normalizedAddress, 48), ",")

    If splitPosition <= 0 Then splitPosition = 48

    firstLine = Trim$(Left$(normalizedAddress, splitPosition))
    secondLine = Trim$(Mid$(normalizedAddress, splitPosition + 1))
End Sub

Private Function BuildEnvelopeAddressLines(addressText As String, postalCode As String, maxLines As Long) As Variant
    Dim result() As String
    ReDim result(1 To maxLines)

    Dim normalizedText As String
    normalizedText = Trim$(addressText)

    If Len(Trim$(postalCode)) > 0 Then
        normalizedText = Replace(normalizedText, ", " & Trim$(postalCode), "", 1, -1, vbTextCompare)
        normalizedText = Replace(normalizedText, Trim$(postalCode), "", 1, -1, vbTextCompare)
    End If

    normalizedText = Replace(normalizedText, vbCrLf, ", ")
    normalizedText = Replace(normalizedText, vbCr, ", ")
    normalizedText = Replace(normalizedText, vbLf, ", ")
    normalizedText = Trim$(normalizedText)

    Dim parts As Variant
    parts = Split(normalizedText, ",")

    Dim lineIndex As Long
    lineIndex = 1

    Dim partIndex As Long
    For partIndex = LBound(parts) To UBound(parts)
        Dim partText As String
        partText = Trim$(CStr(parts(partIndex)))

        If Len(partText) > 0 Then
            If Len(result(lineIndex)) = 0 Then
                result(lineIndex) = partText
            ElseIf Len(result(lineIndex) & ", " & partText) <= 48 Then
                result(lineIndex) = result(lineIndex) & ", " & partText
            ElseIf lineIndex < maxLines Then
                lineIndex = lineIndex + 1
                result(lineIndex) = partText
            Else
                result(lineIndex) = result(lineIndex) & ", " & partText
            End If
        End If
    Next partIndex

    BuildEnvelopeAddressLines = result
End Function

Private Function OnlyDigits(sourceText As String) As String
    Dim charIndex As Long
    For charIndex = 1 To Len(sourceText)
        Dim currentChar As String
        currentChar = Mid$(sourceText, charIndex, 1)

        If currentChar >= "0" And currentChar <= "9" Then OnlyDigits = OnlyDigits & currentChar
    Next charIndex
End Function

Private Function BuildBatchOutgoingNumbersText(batchItems As Collection) As String
    Dim i As Long
    For i = 1 To batchItems.count
        Dim dispatchItem As Variant
        dispatchItem = batchItems(i)

        If i > 1 Then BuildBatchOutgoingNumbersText = BuildBatchOutgoingNumbersText & vbCrLf

        BuildBatchOutgoingNumbersText = BuildBatchOutgoingNumbersText & Trim$(CStr(dispatchItem(DispatchItemColumnLetterNumber)))

        If Len(Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))) > 0 Then
            BuildBatchOutgoingNumbersText = BuildBatchOutgoingNumbersText & " " & t("common.preposition.from", "от") & " " & Trim$(CStr(dispatchItem(DispatchItemColumnLetterDate)))
        End If
    Next i
End Function

Private Function BuildEnvelopeBatchMarker(firstItem As Variant) As String
    BuildEnvelopeBatchMarker = Trim$(CStr(firstItem(DispatchItemColumnBatchId)))

    If Len(BuildEnvelopeBatchMarker) = 0 Then BuildEnvelopeBatchMarker = "dispatch-package-" & Trim$(CStr(firstItem(DispatchItemColumnId)))
End Function

Private Function GetEnvelopeRowsPerPage(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetEnvelopeRowsPerPage = 25
    Case "c5"
        GetEnvelopeRowsPerPage = 22
    Case Else
        GetEnvelopeRowsPerPage = 18
    End Select
End Function

Private Function GetEnvelopeLastColumn(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4", "c5", "dl"
        GetEnvelopeLastColumn = EnvelopeLayoutLastColumnTemplate
    Case Else
        GetEnvelopeLastColumn = EnvelopeLayoutLastColumn
    End Select
End Function

Private Function GetSenderNameCell(envelopeFormatKey As String) As String
    GetSenderNameCell = "C2"
End Function

Private Function GetSenderAddressCell(envelopeFormatKey As String, lineIndex As Long) As String
    Select Case lineIndex
    Case 1
        GetSenderAddressCell = "C3"
    Case 2
        GetSenderAddressCell = "C4"
    Case Else
        GetSenderAddressCell = "C5"
    End Select
End Function

Private Function GetSenderPostalCell(envelopeFormatKey As String) As String
    GetSenderPostalCell = "E7"
End Function

Private Function GetOutgoingCell(envelopeFormatKey As String) As String
    Select Case envelopeFormatKey
    Case "c4"
        GetOutgoingCell = "A9"
    Case "c5"
        GetOutgoingCell = "A9"
    Case Else
        GetOutgoingCell = "A8"
    End Select
End Function

Private Function GetMailTypeMarkCell(envelopeFormatKey As String, lineIndex As Long) As String
    If lineIndex = 1 Then
        GetMailTypeMarkCell = "H2"
    Else
        GetMailTypeMarkCell = "H3"
    End If
End Function

Private Function GetRecipientNameCell(envelopeFormatKey As String) As String
    Select Case envelopeFormatKey
    Case "c4"
        GetRecipientNameCell = "H18"
    Case "c5"
        GetRecipientNameCell = "H15"
    Case Else
        GetRecipientNameCell = "H11"
    End Select
End Function

Private Function GetRecipientAddressCell(envelopeFormatKey As String, lineIndex As Long) As String
    Select Case envelopeFormatKey
    Case "c4"
        If lineIndex = 1 Then GetRecipientAddressCell = "H20"
        If lineIndex = 2 Then GetRecipientAddressCell = "H21"
        If lineIndex >= 3 Then GetRecipientAddressCell = "H22"
    Case "c5"
        If lineIndex = 1 Then GetRecipientAddressCell = "H17"
        If lineIndex = 2 Then GetRecipientAddressCell = "H18"
        If lineIndex >= 3 Then GetRecipientAddressCell = "H19"
    Case Else
        If lineIndex = 1 Then GetRecipientAddressCell = "H13"
        If lineIndex = 2 Then GetRecipientAddressCell = "H14"
        If lineIndex >= 3 Then GetRecipientAddressCell = "H15"
    End Select
End Function

Private Function GetRecipientPostalCell(envelopeFormatKey As String) As String
    Select Case envelopeFormatKey
    Case "c4"
        GetRecipientPostalCell = "H24"
    Case "c5"
        GetRecipientPostalCell = "H21"
    Case Else
        GetRecipientPostalCell = "H17"
    End Select
End Function

Private Function GetLargePostalCell(envelopeFormatKey As String) As String
    Select Case envelopeFormatKey
    Case "c4"
        GetLargePostalCell = "A23"
    Case "c5"
        GetLargePostalCell = "A20"
    Case Else
        GetLargePostalCell = "A16"
    End Select
End Function

Private Function GetBatchMarkerCell(envelopeFormatKey As String) As String
    Select Case envelopeFormatKey
    Case "c4"
        GetBatchMarkerCell = "K25"
    Case "c5"
        GetBatchMarkerCell = "K22"
    Case Else
        GetBatchMarkerCell = "K18"
    End Select
End Function

Private Function GetRecipientTopOffset(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetRecipientTopOffset = 17
    Case "c5"
        GetRecipientTopOffset = 14
    Case Else
        GetRecipientTopOffset = 10
    End Select
End Function

Private Function GetPostalCodeTopOffset(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetPostalCodeTopOffset = 22
    Case "c5"
        GetPostalCodeTopOffset = 19
    Case Else
        GetPostalCodeTopOffset = 15
    End Select
End Function

Private Function GetOutgoingTopOffset(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "dl"
        GetOutgoingTopOffset = 7
    Case Else
        GetOutgoingTopOffset = 8
    End Select
End Function

Private Function GetOutgoingBottomOffset(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "dl"
        GetOutgoingBottomOffset = 9
    Case Else
        GetOutgoingBottomOffset = 12
    End Select
End Function

Private Function GetRecipientBlockHeight(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetRecipientBlockHeight = 4
    Case Else
        GetRecipientBlockHeight = 3
    End Select
End Function

Private Function GetEnvelopeRowHeight(envelopeFormatKey As String) As Double
    Select Case envelopeFormatKey
    Case "c4"
        GetEnvelopeRowHeight = 15.75
    Case "c5"
        GetEnvelopeRowHeight = 16
    Case Else
        GetEnvelopeRowHeight = 14
    End Select
End Function

Private Function GetEnvelopeBaseFontSize(envelopeFormatKey As String) As Integer
    Select Case envelopeFormatKey
    Case "dl"
        GetEnvelopeBaseFontSize = 10
    Case Else
        GetEnvelopeBaseFontSize = 12
    End Select
End Function

Private Function GetEnvelopeSmallFontSize(envelopeFormatKey As String) As Integer
    Select Case envelopeFormatKey
    Case "dl"
        GetEnvelopeSmallFontSize = 8
    Case Else
        GetEnvelopeSmallFontSize = 10
    End Select
End Function

Private Function GetEnvelopePostalFontSize(envelopeFormatKey As String) As Integer
    Select Case envelopeFormatKey
    Case "dl"
        GetEnvelopePostalFontSize = 12
    Case Else
        GetEnvelopePostalFontSize = 14
    End Select
End Function

Private Function GetEnvelopeZoomPercent(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetEnvelopeZoomPercent = 0
    Case Else
        GetEnvelopeZoomPercent = 0
    End Select
End Function

Private Function GetEnvelopePaperSize(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c5"
        GetEnvelopePaperSize = 28
    Case "dl"
        GetEnvelopePaperSize = 27
    Case Else
        GetEnvelopePaperSize = xlPaperA4
    End Select
End Function

Private Function GetEnvelopeLeftMarginCm(envelopeFormatKey As String) As Double
    Select Case envelopeFormatKey
    Case "c4"
        GetEnvelopeLeftMarginCm = 1
    Case "c5"
        GetEnvelopeLeftMarginCm = 1.5
    Case "dl"
        GetEnvelopeLeftMarginCm = 0.5
    Case Else
        GetEnvelopeLeftMarginCm = 0.8
    End Select
End Function

Private Function GetEnvelopeRightMarginCm(envelopeFormatKey As String) As Double
    Select Case envelopeFormatKey
    Case "dl"
        GetEnvelopeRightMarginCm = 1
    Case Else
        GetEnvelopeRightMarginCm = GetEnvelopeLeftMarginCm(envelopeFormatKey)
    End Select
End Function

Private Function GetEnvelopeTopMarginCm(envelopeFormatKey As String) As Double
    GetEnvelopeTopMarginCm = GetEnvelopeLeftMarginCm(envelopeFormatKey)
End Function

Private Function GetEnvelopeBottomMarginCm(envelopeFormatKey As String) As Double
    Select Case envelopeFormatKey
    Case "c5"
        GetEnvelopeBottomMarginCm = 1.2
    Case Else
        GetEnvelopeBottomMarginCm = GetEnvelopeLeftMarginCm(envelopeFormatKey)
    End Select
End Function
