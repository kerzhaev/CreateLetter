Attribute VB_Name = "ModuleEnvelopeLayouts"
' ======================================================================
' Module: ModuleEnvelopeLayouts
' Author: CreateLetter contributors
' Purpose: Prepare hidden workbook layout sheets for C4, C5, and DL envelope batches
' Version: 1.0.0 - 26.04.2026
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

    Dim i As Long
    For i = 1 To dispatchItems.count
        If AppendEnvelopeLayoutRow(dispatchItems(i)) Then
            PrepareEnvelopePrint = PrepareEnvelopePrint + 1
        End If
    Next i

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
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 12)).ClearContents
    End If

    Exit Sub

ClearError:
    Debug.Print "ClearEnvelopeLayoutSheetData error: " & Err.description
End Sub

Private Function AppendEnvelopeLayoutRow(dispatchItem As Variant) As Boolean
    Dim sheetName As String
    sheetName = ResolveEnvelopeLayoutSheetName(CStr(dispatchItem(DispatchItemColumnEnvelopeFormatKey)))
    If Len(sheetName) = 0 Then Exit Function

    On Error GoTo AppendError

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim NextRow As Long
    NextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2

    ws.Cells(NextRow, 1).value = CStr(dispatchItem(DispatchItemColumnId))
    ws.Cells(NextRow, 2).value = CStr(dispatchItem(DispatchItemColumnLetterNumber))
    ws.Cells(NextRow, 3).value = CStr(dispatchItem(DispatchItemColumnAddressee))
    ws.Cells(NextRow, 4).value = CStr(dispatchItem(DispatchItemColumnAddressLine))
    ws.Cells(NextRow, 5).value = CStr(dispatchItem(DispatchItemColumnPostalCode))
    ws.Cells(NextRow, 6).value = CStr(dispatchItem(DispatchItemColumnSenderName))
    ws.Cells(NextRow, 7).value = DispatchRepositoryGetSenderPostalCode(CStr(dispatchItem(DispatchItemColumnSenderName)))
    ws.Cells(NextRow, 8).value = CStr(dispatchItem(DispatchItemColumnMailType))
    ws.Cells(NextRow, 9).value = CStr(dispatchItem(DispatchItemColumnMass))
    ws.Cells(NextRow, 10).value = CStr(dispatchItem(DispatchItemColumnDeclaredValue))
    ws.Cells(NextRow, 11).value = CStr(dispatchItem(DispatchItemColumnComment))
    ws.Cells(NextRow, 12).value = Format$(Now, "dd.mm.yyyy hh:nn:ss")

    AppendEnvelopeLayoutRow = True
    Exit Function

AppendError:
    Debug.Print "AppendEnvelopeLayoutRow error: " & Err.description
    AppendEnvelopeLayoutRow = False
End Function