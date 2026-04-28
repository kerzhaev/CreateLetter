Attribute VB_Name = "ModuleEnvelopeLayouts"
' ======================================================================
' Module: ModuleEnvelopeLayouts
' Author: CreateLetter contributors
' Purpose: Prepare printable workbook layout sheets for grouped C4, C5, and DL envelope batches
' Version: 1.2.0 - 28.04.2026
' ======================================================================

Option Explicit

Private Const EnvelopeLayoutSheetC4 As String = "DispatchLayout_C4"
Private Const EnvelopeLayoutSheetC5 As String = "DispatchLayout_C5"
Private Const EnvelopeLayoutSheetDL As String = "DispatchLayout_DL"
Private Const EnvelopeLayoutFirstColumn As Long = 1
Private Const EnvelopeLayoutLastColumn As Long = 6

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

    Dim firstVisibleSheet As Worksheet
    Dim batchKey As Variant
    For Each batchKey In groupedBatches.keys
        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)
        If Not batchItems Is Nothing Then
            If batchItems.count > 0 Then
                If AppendEnvelopeLayoutPage(batchItems, firstVisibleSheet) Then
                    PrepareEnvelopePrint = PrepareEnvelopePrint + 1
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

    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ws.ResetAllPageBreaks
    ws.Cells.Font.Name = "Times New Roman"
    ws.Cells.Font.Size = 12
    ws.Cells.VerticalAlignment = xlTop
    ws.Visible = xlSheetVeryHidden

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
        If Len(batchKey) = 0 Then batchKey = CStr(dispatchItem(DispatchItemColumnId))

        If Not groupedBatches.Exists(batchKey) Then groupedBatches.Add batchKey, New Collection

        Dim batchItems As Collection
        Set batchItems = groupedBatches.item(batchKey)
        batchItems.Add dispatchItem
    Next i

    Set GroupDispatchItemsByBatch = groupedBatches
End Function

Private Function AppendEnvelopeLayoutPage(batchItems As Collection, ByRef firstVisibleSheet As Worksheet) As Boolean
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

    ConfigureEnvelopeLayoutGrid ws, envelopeFormatKey
    RenderEnvelopeLayoutBlock ws, topRow, envelopeFormatKey, batchItems
    ConfigureEnvelopePageSettings ws, envelopeFormatKey

    If firstVisibleSheet Is Nothing Then Set firstVisibleSheet = ws

    AppendEnvelopeLayoutPage = True
    Exit Function

AppendError:
    Debug.Print "AppendEnvelopeLayoutPage error: " & Err.description
    AppendEnvelopeLayoutPage = False
End Function

Private Function GetNextEnvelopeTopRow(ws As Worksheet, envelopeFormatKey As String) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetNextEnvelopeTopRow = 1
    Else
        GetNextEnvelopeTopRow = ws.Cells(ws.Rows.count, EnvelopeLayoutFirstColumn).End(xlUp).Row + 1
    End If
End Function

Private Sub ConfigureEnvelopeLayoutGrid(ws As Worksheet, envelopeFormatKey As String)
    Dim colIndex As Long
    For colIndex = EnvelopeLayoutFirstColumn To EnvelopeLayoutLastColumn
        ws.Columns(colIndex).ColumnWidth = 12
    Next colIndex

    Select Case envelopeFormatKey
    Case "c4"
        ws.Columns("A").ColumnWidth = 12
        ws.Columns("B").ColumnWidth = 14
        ws.Columns("C").ColumnWidth = 14
        ws.Columns("D").ColumnWidth = 18
        ws.Columns("E").ColumnWidth = 18
        ws.Columns("F").ColumnWidth = 18
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
    Set blockRange = ws.Range(ws.Cells(topRow, EnvelopeLayoutFirstColumn), ws.Cells(topRow + rowsPerPage - 1, EnvelopeLayoutLastColumn))
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

Private Sub ConfigureEnvelopePageSettings(ws As Worksheet, envelopeFormatKey As String)
    Dim lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    If lastRow < 1 Then lastRow = 1

    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, EnvelopeLayoutFirstColumn), ws.Cells(lastRow, EnvelopeLayoutLastColumn)).Address

    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(GetEnvelopeMarginCm(envelopeFormatKey))
        .RightMargin = Application.CentimetersToPoints(GetEnvelopeMarginCm(envelopeFormatKey))
        .TopMargin = Application.CentimetersToPoints(GetEnvelopeMarginCm(envelopeFormatKey))
        .BottomMargin = Application.CentimetersToPoints(GetEnvelopeMarginCm(envelopeFormatKey))
        .CenterHorizontally = True
        .CenterVertically = True
    End With
End Sub

Private Sub FinalizeEnvelopeLayoutSheet(sheetName As String)
    On Error GoTo FinalizeError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

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
        GetEnvelopeRowsPerPage = 12
    Case "c5"
        GetEnvelopeRowsPerPage = 10
    Case Else
        GetEnvelopeRowsPerPage = 8
    End Select
End Function

Private Function GetRecipientTopOffset(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetRecipientTopOffset = 4
    Case "c5"
        GetRecipientTopOffset = 3
    Case Else
        GetRecipientTopOffset = 2
    End Select
End Function

Private Function GetPostalCodeTopOffset(envelopeFormatKey As String) As Long
    Select Case envelopeFormatKey
    Case "c4"
        GetPostalCodeTopOffset = 10
    Case "c5"
        GetPostalCodeTopOffset = 8
    Case Else
        GetPostalCodeTopOffset = 6
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
        GetEnvelopeRowHeight = 18
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

Private Function GetEnvelopeMarginCm(envelopeFormatKey As String) As Double
    Select Case envelopeFormatKey
    Case "dl"
        GetEnvelopeMarginCm = 0.6
    Case Else
        GetEnvelopeMarginCm = 0.8
    End Select
End Function
