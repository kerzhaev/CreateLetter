Attribute VB_Name = "ModulePostalRegistryPrint"
' ======================================================================
' Module: ModulePostalRegistryPrint
' Author: CreateLetter contributors
' Purpose: Build a printable postal registry sheet from the internal dispatch registry
' Version: 1.0.1 - 27.04.2026
' ======================================================================

Option Explicit

Private Const PostalRegistryPrintSheetName As String = "PostalRegistryPrint"
Private Const PostalRegistryHeaderRow As Long = 7
Private Const PostalRegistryFirstDataRow As Long = 8

Public Function BuildPostalRegistryPrintSheet() As Long
    On Error GoTo BuildError

    Dim registryTable As ListObject
    Set registryTable = GetPostalRegistrySourceTable()

    Dim printSheet As Worksheet
    Set printSheet = GetOrCreatePostalRegistryPrintSheet()

    PreparePostalRegistryPrintSheet printSheet

    If registryTable.DataBodyRange Is Nothing Then Exit Function

    Dim registryData As Variant
    registryData = registryTable.DataBodyRange.Value2

    WritePostalRegistryHeader printSheet, registryData

    Dim writtenRows As Long
    WritePostalRegistryTable printSheet, registryData, writtenRows
    WritePostalRegistryFooter printSheet, PostalRegistryFirstDataRow + writtenRows + 2
    ConfigurePostalRegistryPage printSheet

    BuildPostalRegistryPrintSheet = writtenRows
    printSheet.Activate
    Exit Function

BuildError:
    Debug.Print "BuildPostalRegistryPrintSheet error: " & Err.Description
    BuildPostalRegistryPrintSheet = 0
End Function

Private Function GetPostalRegistrySourceTable() As ListObject
    Set GetPostalRegistrySourceTable = ThisWorkbook.Worksheets("DispatchRegistry").ListObjects.Item(DispatchRegistryTableName)
End Function

Private Function GetOrCreatePostalRegistryPrintSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePostalRegistryPrintSheet = ThisWorkbook.Worksheets(PostalRegistryPrintSheetName)
    On Error GoTo 0

    If GetOrCreatePostalRegistryPrintSheet Is Nothing Then
        Set GetOrCreatePostalRegistryPrintSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreatePostalRegistryPrintSheet.Name = PostalRegistryPrintSheetName
    End If

    GetOrCreatePostalRegistryPrintSheet.Visible = xlSheetVisible
End Function

Private Sub PreparePostalRegistryPrintSheet(targetSheet As Worksheet)
    targetSheet.Cells.Clear
    targetSheet.Cells.Font.Name = "Times New Roman"
    targetSheet.Cells.Font.Size = 12

    targetSheet.Columns("A").ColumnWidth = 5
    targetSheet.Columns("B").ColumnWidth = 10
    targetSheet.Columns("C").ColumnWidth = 30
    targetSheet.Columns("D").ColumnWidth = 28
    targetSheet.Columns("E").ColumnWidth = 21
    targetSheet.Columns("F").ColumnWidth = 16
End Sub

Private Sub WritePostalRegistryHeader(targetSheet As Worksheet, registryData As Variant)
    Dim registryNumber As String
    registryNumber = FirstNonEmptyRegistryValue(registryData, DispatchRegistryColumnRegistryNumber)

    Dim registryDate As String
    registryDate = FirstNonEmptyRegistryValue(registryData, DispatchRegistryColumnRegistryDate)

    Dim senderName As String
    senderName = FirstNonEmptyRegistryValue(registryData, DispatchRegistryColumnSenderName)

    With targetSheet
        .Range("A1:F1").Merge
        .Range("A1").Value = t("postal.registry.print.registry_prefix", "Registry No. ") & registryNumber
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2:F2").Merge
        .Range("A2").Value = t("postal.registry.print.submitted_to", "Correspondence submitted to ") & GetPostalRegistryPostOfficeName()
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A2").Font.Size = 13

        .Range("A3:F4").Merge
        .Range("A3").Value = t("postal.registry.print.sender_prefix", "Sender: ") & senderName
        .Range("A3").HorizontalAlignment = xlCenter
        .Range("A3").VerticalAlignment = xlCenter
        .Range("A3").WrapText = True
        .Range("A3").Font.Size = 13

        .Range("A5:F5").Merge
        .Range("A5").Value = registryDate
        .Range("A5").HorizontalAlignment = xlCenter
        .Range("A5").Font.Size = 13
    End With
End Sub

Private Sub WritePostalRegistryTable(targetSheet As Worksheet, registryData As Variant, ByRef writtenRows As Long)
    With targetSheet
        .Cells(PostalRegistryHeaderRow, 1).Value = t("postal.registry.print.column.number", "No.")
        .Cells(PostalRegistryHeaderRow, 2).Value = t("postal.registry.print.column.index", "Index")
        .Cells(PostalRegistryHeaderRow, 3).Value = t("postal.registry.print.column.destination", "Destination")
        .Cells(PostalRegistryHeaderRow, 4).Value = t("postal.registry.print.column.addressee", "Addressee")
        .Cells(PostalRegistryHeaderRow, 5).Value = t("postal.registry.print.column.letter_number", "Letter No.")
        .Cells(PostalRegistryHeaderRow, 6).Value = t("postal.registry.print.column.note", "Note")

        With .Range(.Cells(PostalRegistryHeaderRow, 1), .Cells(PostalRegistryHeaderRow, 6))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    End With

    Dim sourceRow As Long
    For sourceRow = 1 To UBound(registryData, 1)
        Dim targetRow As Long
        targetRow = PostalRegistryFirstDataRow + writtenRows

        With targetSheet
            .Cells(targetRow, 1).Value = writtenRows + 1
            .Cells(targetRow, 2).Value = CStr(registryData(sourceRow, DispatchRegistryColumnPostalCode))
            .Cells(targetRow, 3).Value = CStr(registryData(sourceRow, DispatchRegistryColumnAddressLine))
            .Cells(targetRow, 4).Value = CStr(registryData(sourceRow, DispatchRegistryColumnAddressee))
            .Cells(targetRow, 5).Value = BuildPostalRegistryOutgoingCell(CStr(registryData(sourceRow, DispatchRegistryColumnOutgoingNumbers)))
            .Cells(targetRow, 6).Value = UCase$(CStr(registryData(sourceRow, DispatchRegistryColumnMailType)))

            With .Range(.Cells(targetRow, 1), .Cells(targetRow, 6))
                .WrapText = True
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With

            .Cells(targetRow, 1).HorizontalAlignment = xlCenter
            .Cells(targetRow, 2).HorizontalAlignment = xlCenter
            .Cells(targetRow, 6).HorizontalAlignment = xlCenter
            .Rows(targetRow).RowHeight = EstimatePostalRegistryRowHeight(CStr(registryData(sourceRow, DispatchRegistryColumnOutgoingNumbers)))
        End With

        writtenRows = writtenRows + 1
    Next sourceRow
End Sub

Private Sub WritePostalRegistryFooter(targetSheet As Worksheet, footerStartRow As Long)
    With targetSheet
        .Cells(footerStartRow, 1).Value = t("postal.registry.print.footer.total", "TOTAL")
        .Cells(footerStartRow, 2).Value = CountPostalRegistryPackages()
        .Cells(footerStartRow, 3).Value = t("postal.registry.print.footer.package", "package.")
        .Range(.Cells(footerStartRow, 1), .Cells(footerStartRow, 3)).Font.Bold = True

        .Cells(footerStartRow + 2, 1).Value = t("postal.registry.print.footer.sender_signature", "Sender signature:")
        .Range(.Cells(footerStartRow + 2, 2), .Cells(footerStartRow + 2, 4)).Merge
        .Cells(footerStartRow + 2, 2).Value = "________________________"

        .Cells(footerStartRow + 4, 1).Value = t("postal.registry.print.footer.stamp", "Stamp")

        .Cells(footerStartRow + 6, 1).Value = t("postal.registry.print.footer.accepted_by_registry", "Accepted by this registry:")
        .Range(.Cells(footerStartRow + 6, 3), .Cells(footerStartRow + 6, 4)).Merge
        .Cells(footerStartRow + 6, 3).Value = t("postal.registry.print.footer.documents", "____ documents.")

        .Cells(footerStartRow + 8, 1).Value = t("postal.registry.print.footer.stamp", "Stamp")

        .Cells(footerStartRow + 10, 1).Value = t("postal.registry.print.footer.quote_open", """")
        .Cells(footerStartRow + 10, 2).Value = t("postal.registry.print.footer.quote_close", """")
        .Cells(footerStartRow + 10, 3).Value = "202"
        .Cells(footerStartRow + 10, 4).Value = t("postal.registry.print.footer.year_word", "year")

        .Cells(footerStartRow + 12, 1).Value = t("postal.registry.print.footer.receiver_signature", "Receiver signature")
        .Range(.Cells(footerStartRow + 12, 2), .Cells(footerStartRow + 12, 4)).Merge
        .Cells(footerStartRow + 12, 2).Value = "________________________"
    End With
End Sub

Private Sub ConfigurePostalRegistryPage(targetSheet As Worksheet)
    With targetSheet.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1.4)
        .RightMargin = Application.CentimetersToPoints(1.4)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
    End With
End Sub

Private Function FirstNonEmptyRegistryValue(registryData As Variant, columnIndex As Long) As String
    Dim rowIndex As Long
    For rowIndex = 1 To UBound(registryData, 1)
        If Len(Trim$(CStr(registryData(rowIndex, columnIndex)))) > 0 Then
            FirstNonEmptyRegistryValue = CStr(registryData(rowIndex, columnIndex))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function GetPostalRegistryPostOfficeName() As String
    GetPostalRegistryPostOfficeName = GetSetting("CreateLetter", "PostalRegistry", "PostOfficeName", "")
    If Len(Trim$(GetPostalRegistryPostOfficeName)) = 0 Then
        GetPostalRegistryPostOfficeName = t("postal.registry.print.default_post_office", "post office")
    End If
End Function

Private Function BuildPostalRegistryOutgoingCell(outgoingNumbersText As String) As String
    Dim normalizedText As String
    normalizedText = Replace(outgoingNumbersText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(normalizedText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(CStr(lines(i)))
        If Len(lineText) > 0 Then
            If Len(BuildPostalRegistryOutgoingCell) > 0 Then
                BuildPostalRegistryOutgoingCell = BuildPostalRegistryOutgoingCell & vbCrLf
            End If
            BuildPostalRegistryOutgoingCell = BuildPostalRegistryOutgoingCell & t("postal.registry.print.outgoing_prefix", "Out. No. ") & AddRegistryDateSuffix(lineText)
        End If
    Next i
End Function

Private Function AddRegistryDateSuffix(lineText As String) As String
    Dim dateMarker As String
    dateMarker = " " & t("common.preposition.from", "dated") & " "

    If InStr(1, lineText, dateMarker, vbTextCompare) = 0 Then
        AddRegistryDateSuffix = lineText
        Exit Function
    End If

    Dim yearSuffix As String
    yearSuffix = " " & t("postal.registry.print.year_suffix", "yr.")

    If Right$(Trim$(lineText), Len(Trim$(yearSuffix))) = Trim$(yearSuffix) Then
        AddRegistryDateSuffix = lineText
    Else
        AddRegistryDateSuffix = lineText & yearSuffix
    End If
End Function

Private Function EstimatePostalRegistryRowHeight(outgoingNumbersText As String) As Double
    Dim lineCount As Long
    lineCount = UBound(Split(Replace(outgoingNumbersText, vbCrLf, vbLf), vbLf)) + 1
    If lineCount < 1 Then lineCount = 1

    EstimatePostalRegistryRowHeight = 48 + ((lineCount - 1) * 22)
End Function

Private Function CountPostalRegistryPackages() As Long
    Dim registryTable As ListObject
    Set registryTable = GetPostalRegistrySourceTable()

    If registryTable.DataBodyRange Is Nothing Then Exit Function
    CountPostalRegistryPackages = registryTable.DataBodyRange.Rows.Count
End Function
