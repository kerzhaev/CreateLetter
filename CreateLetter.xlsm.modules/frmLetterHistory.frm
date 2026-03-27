VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLetterHistory 
   Caption         =   "Letter History"
   ClientHeight    =   11715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13515
   OleObjectBlob   =   "frmLetterHistory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLetterHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================
' Form: frmLetterHistory v1.2.2 - Thin-shell history UI
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Date: 27.03.2026
' Purpose: History of sent letters with thin-shell UI, navigation, filtering, status updates, and schema-safe bindings
' Updates v1.2.2:
' - Replaced hardcoded history field indexes with shared ModuleMain enums
' - Kept navigation, export, and status update flow aligned with named letter columns
' - Preserved Russian/European date formatting and record navigation workflow
' ======================================================================
Option Explicit

Private allLettersData As Collection
Private filteredData As Collection

' ===============================================================================
' FORM INITIALIZATION - IMPROVED VERSION
' ===============================================================================
Private Sub UserForm_Initialize()
    Set allLettersData = New Collection
    Set filteredData = New Collection
    
    ApplyFormSettings
    ApplyEnglishStaticCaptions
    ApplyElementStyles
    ConfigureDateFieldRussianFormat  ' NEW: Russian/European date format
    ConfigureCompactSumField         ' NEW: Compact sum field
    
    LoadAllLettersData
    ShowAllLettersOnInit
    InitializeControlValues
    
        Debug.Print "Letter history form initialized v1.2.2 with thin-shell helpers"
End Sub

Private Sub ConfigureDateFieldRussianFormat()
    ' FIXED FUNCTION: Date field configuration without non-existent properties
    On Error Resume Next
    
    If Not dtpReturnDate Is Nothing Then
        With dtpReturnDate
            ' FIXED: Setting the date in Russian/European format via Format()
            .Value = Format(Date, "dd.mm.yyyy")
            
            ' REMOVED: .Format = 2 (not supported by TextBox)
            ' REMOVED: .CustomFormat = "dd.mm.yyyy" (not supported by TextBox)
            
            .ControlTipText = "Document return date (dd.mm.yyyy)"
        End With
        Debug.Print "Return date field configured in Russian/European format"
    End If
    
    On Error GoTo 0
End Sub



Private Sub ConfigureCompactSumField()
    ' FIXED FUNCTION: Compact sum field configuration without EnterKeyBehaviour
    On Error Resume Next
    
    If Not txtSumDocument Is Nothing Then
        With txtSumDocument
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Multiline = False              ' Single-line field
            .WordWrap = False               ' No word wrap
            .ScrollBars = 0                 ' No scrollbars (0 = fmScrollBarsNone)
            ' REMOVED: .EnterKeyBehaviour = False (not supported)
            .Height = 24                    ' Fixed height
            .ControlTipText = "Document sum in rubles (numbers only or brief comment)"
        End With
        Debug.Print "Document sum field configured in compact mode"
    End If
    
    On Error GoTo 0
End Sub


Private Sub ApplyFormSettings()
    With Me
        .Caption = "Letter History v1.2.2"
        .backColor = RGB(250, 250, 250)
    End With
End Sub

Private Sub ApplyEnglishStaticCaptions()
    On Error Resume Next

    SetHistoryControlCaption "frameSearch", "Search"
    SetHistoryControlCaption "frameLettersList", "Letter list"
    SetHistoryControlCaption "frameStatusEdit", "Status update"
    SetHistoryControlCaption "frameActions", "Actions"
    SetHistoryControlCaption "Label1", "Search letters by delivery history"
    SetHistoryControlCaption "lblSearchLabel", "Search status"
    SetHistoryControlCaption "lblSearchInfo", "Ready"
    SetHistoryControlCaption "lblDateLabel", "Return date"
    SetHistoryControlCaption "lblSumLabel", "Amount"
    SetHistoryControlCaption "btnUpdateStatus", "Update status"
    SetHistoryControlCaption "btnClose", "Close"
    SetHistoryControlCaption "btnRefresh", "Refresh data"
    SetHistoryControlCaption "btnClearSearch", "Clear search"
    SetHistoryControlCaption "btnExportToExcel", "Export to Excel"
    SetHistoryControlCaption "btnNavigateToRecord", "Go to record"
    SetHistoryControlCaption "btnSearchHelp", "Search help"
    SetHistoryControlCaption "chkReceived", "Received back"

    On Error GoTo 0
End Sub

Private Sub SetHistoryControlCaption(controlName As String, captionText As String)
    On Error Resume Next

    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then
        ctrl.Caption = captionText
    End If

    On Error GoTo 0
End Sub






Private Sub InitializeControlValues()
    On Error Resume Next
    
    Me.Controls("txtHistorySearch").Value = ""
    Me.Controls("txtSumDocument").Value = ""
    Me.Controls("chkReceived").Value = False
    
    ' FIXED: Formatting date when setting value
    Me.Controls("dtpReturnDate").Value = Format(Date, "dd.mm.yyyy")
    
    If Not lstLetterHistory Is Nothing Then
        lstLetterHistory.Clear
    End If
    
    Debug.Print "Letter history form initialized correctly"
    
    On Error GoTo 0
End Sub


' ===============================================================================
' DATA LOADING AND PROCESSING (NO CHANGES)
' ===============================================================================
Private Sub LoadAllLettersData()
    Set allLettersData = LoadLetterHistoryData()
    
    If allLettersData Is Nothing Or allLettersData.count = 0 Then
        UpdateSearchInfo "No data found in worksheet 'Letters'"
        Exit Sub
    End If
    
    UpdateSearchInfo "Letters loaded: " & allLettersData.count
End Sub



Private Sub ShowAllLettersOnInit()
    If lstLetterHistory Is Nothing Then Exit Sub
    
    Set filteredData = FilterLetterHistoryRecords(allLettersData, "")
    BindHistoryList filteredData
    
    UpdateSearchInfo "Showing all letters: " & allLettersData.count
End Sub

' ===============================================================================
' SEARCH AND FILTERING (NO CHANGES)
' ===============================================================================
Private Sub txtHistorySearch_Change()
    If txtHistorySearch Is Nothing Then Exit Sub
    
    Dim searchText As String
    searchText = Trim(txtHistorySearch.Value)
    
    ' DEBUG: Outputting search information
    If IsNumeric(searchText) And Len(searchText) > 2 Then
        Debug.Print "=== SEARCH DEBUG ==="
        Debug.Print "Searching for number: " & searchText
        Debug.Print "Total records to search: " & allLettersData.count
        
        ' Show first few records for testing
        Dim i As Integer
        For i = 1 To WorksheetFunction.Min(3, allLettersData.count)
            Dim parts() As String
            parts = Split(allLettersData(i), "|")
            If UBound(parts) >= 4 Then
                Debug.Print "Record " & i & ", sum column: '" & parts(HistoryPartDocumentSum) & "'"
            End If
        Next i
        Debug.Print "=================="
    End If
    
    If Len(searchText) = 0 Then
        ShowAllLettersOnInit
    Else
        Set filteredData = FilterLetterHistoryRecords(allLettersData, searchText)
        BindHistoryList filteredData
        
        If IsNumeric(searchText) And Len(searchText) > 2 Then
            UpdateSearchInfo "Searching for number " & searchText & " in document amounts..."
        End If
    End If
End Sub

Private Sub UpdateSearchInfo(message As String)
    On Error Resume Next
    If Not lblSearchInfo Is Nothing Then
        lblSearchInfo.Caption = message
    End If
    On Error GoTo 0
End Sub

' ===============================================================================
' CONTROL EVENTS - REVISED
' ===============================================================================
Private Sub lstLetterHistory_Click()
    If lstLetterHistory Is Nothing Then Exit Sub
    If lstLetterHistory.ListIndex < 0 Then Exit Sub
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.ListIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterData As String
        letterData = filteredData(selectedIndex)
        
        Dim parts() As String
        If TryParseLetterHistoryRecord(letterData, parts) Then
            On Error Resume Next
            
            If Not txtSumDocument Is Nothing Then
                txtSumDocument.Value = parts(HistoryPartDocumentSum)
            End If
            
            ParseReturnStatus parts(HistoryPartReturnStatus)
            
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub lstLetterHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' NEW EVENT: Double click to navigate to Excel record
    NavigateToSelectedRecord
End Sub

Private Sub btnNavigateToRecord_Click()
    ' NEW BUTTON: Navigate to selected record
    NavigateToSelectedRecord
End Sub

Private Sub NavigateToSelectedRecord()
    ' IMPROVED FUNCTION: Navigation to record retaining focus on the form
    On Error GoTo NavigateError
    
    If lstLetterHistory Is Nothing Then Exit Sub
    If lstLetterHistory.ListIndex < 0 Then
        MsgBox "Select a letter to navigate to the record.", vbExclamation, "Go to record"
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.ListIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterData As String
        letterData = filteredData(selectedIndex)
        
        Dim parts() As String
        If TryParseLetterHistoryRecord(letterData, parts) Then
            Dim rowNumber As Long
            rowNumber = CLng(parts(HistoryPartRowNumber))
            
            ' Getting "Letters" sheet
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets("Letters")
            
            If ws Is Nothing Then
                MsgBox "Worksheet 'Letters' not found.", vbCritical, "Navigation error"
                Exit Sub
            End If
            
            ' IMPROVED: Activate Excel and jump to record
            Application.Visible = True
            ws.Activate
            ws.Cells(rowNumber, LetterColumnAddressee).Select
            
            ' Highlight record row
            With ws.Rows(rowNumber).Interior
                .Color = RGB(255, 255, 0)  ' Yellow highlight
                .Pattern = xlSolid
            End With
            
            ' NEW: Show info in Excel status bar
            Application.StatusBar = "Selected record: " & parts(HistoryPartAddressee) & " | " & parts(HistoryPartOutgoingNumber) & " | " & parts(HistoryPartOutgoingDate)
            
            ' Remove highlight after 5 seconds
            Application.OnTime Now + TimeValue("00:00:05"), "ClearHighlight"
            
            ' IMPORTANT: DO NOT show MsgBox, to not interfere with Excel workflow
            Debug.Print "Navigated to record, row: " & rowNumber
            
            ' NEW: Return focus to history form after 1 second
            Application.OnTime Now + TimeValue("00:00:01"), "RestoreFocusToHistory"
        End If
    End If
    
    Exit Sub
    
NavigateError:
    MsgBox "Error navigating to record: " & Err.description, vbCritical, "Navigation error"
End Sub


Private Sub ParseReturnStatus(returnStatus As String)
    On Error Resume Next
    
    If InStr(UCase(returnStatus), "RECEIVED") > 0 And InStr(UCase(returnStatus), "NOT RECEIVED") = 0 Then
        If Not chkReceived Is Nothing Then
            chkReceived.Value = True
        End If
        
        Dim dateString As String
        dateString = ExtractDateFromReturnStatus(returnStatus)
        
        If IsDate(dateString) And Not dtpReturnDate Is Nothing Then
            ' FIXED: Formatting date in Russian/European format
            dtpReturnDate.Value = Format(CDate(dateString), "dd.mm.yyyy")
        ElseIf Not dtpReturnDate Is Nothing Then
            dtpReturnDate.Value = Format(Date, "dd.mm.yyyy")
        End If
    Else
        If Not chkReceived Is Nothing Then
            chkReceived.Value = False
        End If
        If Not dtpReturnDate Is Nothing Then
            ' FIXED: Formatting date in Russian/European format
            dtpReturnDate.Value = Format(Date, "dd.mm.yyyy")
        End If
    End If
    
    On Error GoTo 0
End Sub


Private Function ExtractDateFromReturnStatus(returnStatus As String) As String
    Dim result As String
    result = ""
    
    Dim parts() As String
    parts = Split(returnStatus, " ")
    
    Dim j As Integer
    For j = 0 To UBound(parts)
        If IsDate(parts(j)) Then
            result = parts(j)
            Exit Function
        End If
    Next j
    
    ExtractDateFromReturnStatus = result
End Function

' ===============================================================================
' ACTION BUTTONS (NO CHANGES)
' ===============================================================================
Private Sub btnClearSearch_Click()
    On Error Resume Next
    
    Dim originalCaption As String
    originalCaption = Me.Controls("btnClearSearch").Caption
    Me.Controls("btnClearSearch").Caption = "Clearing..."
    Me.Controls("btnClearSearch").Enabled = False
    
    DoEvents
    
    Me.Controls("txtHistorySearch").Value = ""
    ClearAllHistoryFields
    ShowAllLettersOnInit
    
    Me.Controls("btnClearSearch").Caption = originalCaption
    Me.Controls("btnClearSearch").Enabled = True
    Me.Controls("txtHistorySearch").SetFocus
    
    Debug.Print "Full clear of letter history form executed"
    
    On Error GoTo 0
End Sub

Private Sub ClearAllHistoryFields()
    On Error Resume Next
    
    Me.Controls("txtSumDocument").Value = ""
    Me.Controls("chkReceived").Value = False
    
    ' FIXED: Format date upon clearing
    Me.Controls("dtpReturnDate").Value = Format(Date, "dd.mm.yyyy")
    
    SetControlBackColor "txtSumDocument", RGB(255, 255, 255)
    
    Debug.Print "All letter history fields cleared"
    
    On Error GoTo 0
End Sub


Private Sub SetControlBackColor(controlName As String, backColor As Long)
    On Error Resume Next
    
    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then
        ctrl.backColor = backColor
    End If
    
    On Error GoTo 0
End Sub

Private Sub btnUpdateStatus_Click()
    If lstLetterHistory Is Nothing Then Exit Sub
    If lstLetterHistory.ListIndex < 0 Then
        MsgBox "Select a letter to update the status.", vbExclamation
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.ListIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterData As String
        letterData = filteredData(selectedIndex)
        
        Dim parts() As String
        If TryParseLetterHistoryRecord(letterData, parts) Then
            Dim rowNumber As Long
            rowNumber = CLng(parts(HistoryPartRowNumber))
            Dim returnStatus As String
            returnStatus = BuildLetterReturnStatus((Not chkReceived Is Nothing And chkReceived.Value), ControlValueOrDefault("dtpReturnDate"))
            
            UpdateLetterHistoryRow rowNumber, ControlValueOrDefault("txtSumDocument"), returnStatus
            
            LoadAllLettersData
            txtHistorySearch_Change
            
            MsgBox "Letter status updated successfully.", vbInformation
        End If
    End If
End Sub

Private Sub BindHistoryList(records As Collection)
    If lstLetterHistory Is Nothing Then Exit Sub
    
    lstLetterHistory.Clear
    
    Dim i As Long
    For i = 1 To records.count
        lstLetterHistory.AddItem FormatLetterHistoryDisplay(CStr(records(i)))
    Next i
    
    UpdateSearchInfo "Letters found: " & records.count & " of " & allLettersData.count
End Sub

Private Function ControlValueOrDefault(controlName As String, Optional defaultValue As String = "") As String
    On Error Resume Next
    ControlValueOrDefault = Trim(CStr(Me.Controls(controlName).Value))
    If Err.number <> 0 Then
        ControlValueOrDefault = defaultValue
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub btnRefresh_Click()
    LoadAllLettersData
    txtHistorySearch_Change
    MsgBox "Data refreshed.", vbInformation
End Sub

Private Sub btnExportToExcel_Click()
    If filteredData.count = 0 Then
        MsgBox "No data to export.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ExportError
    
    Dim exportWb As Workbook
    Dim exportWs As Worksheet
    Set exportWb = Workbooks.Add
    Set exportWs = exportWb.Worksheets(1)
    
    With exportWs
        .Cells(1, LetterColumnAddressee).Value = "Addressee"
        .Cells(1, LetterColumnOutgoingNumber).Value = "Outgoing Number"
        .Cells(1, LetterColumnOutgoingDate).Value = "Outgoing Date"
        .Cells(1, LetterColumnAttachmentText).Value = "Attachment Name"
        .Cells(1, LetterColumnDocumentSum).Value = "Document Sum"
        .Cells(1, LetterColumnReturnStatus).Value = "Return Mark"
        .Cells(1, LetterColumnExecutor).Value = "Executor Name"
        .Cells(1, LetterColumnDocumentType).Value = "Send Type"
        
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.ColorIndex = 15
            .Font.ColorIndex = 1
        End With
    End With
    
    Dim i As Integer
    For i = 1 To filteredData.count
        Dim letterData As String
        letterData = filteredData(i)
        
        Dim parts() As String
        parts = Split(letterData, "|")
        
        If UBound(parts) >= HistoryPartDocumentType Then
            exportWs.Cells(i + 1, LetterColumnAddressee).Value = parts(HistoryPartAddressee)
            exportWs.Cells(i + 1, LetterColumnOutgoingNumber).Value = parts(HistoryPartOutgoingNumber)
            exportWs.Cells(i + 1, LetterColumnOutgoingDate).Value = parts(HistoryPartOutgoingDate)
            exportWs.Cells(i + 1, LetterColumnAttachmentText).Value = parts(HistoryPartAttachmentText)
            exportWs.Cells(i + 1, LetterColumnDocumentSum).Value = parts(HistoryPartDocumentSum)
            exportWs.Cells(i + 1, LetterColumnReturnStatus).Value = parts(HistoryPartReturnStatus)
            exportWs.Cells(i + 1, LetterColumnExecutor).Value = parts(HistoryPartExecutor)
            exportWs.Cells(i + 1, LetterColumnDocumentType).Value = parts(HistoryPartDocumentType)
        End If
    Next i
    
    exportWs.Columns("A:H").AutoFit
    exportWs.Name = "Letters history " & Format(Date, "dd.mm.yyyy")
    exportWb.Application.Visible = True
    
    MsgBox "Export completed." & vbCrLf & "Records exported: " & filteredData.count, vbInformation, "Data export"
    Exit Sub
    
ExportError:
    MsgBox "Export error: " & Err.description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set allLettersData = Nothing
    Set filteredData = Nothing
    
    Debug.Print "Letter history form closed, memory freed"
End Sub


Private Sub dtpReturnDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' NEW FUNCTION: Validating entered date
    On Error Resume Next
    
    If Not dtpReturnDate Is Nothing Then
        Dim inputText As String
        inputText = Trim(dtpReturnDate.Value)
        
        ' If field is not empty, validate date
        If Len(inputText) > 0 Then
            If IsDate(inputText) Then
                ' Format valid date in Russian/European format
                dtpReturnDate.Value = Format(CDate(inputText), "dd.mm.yyyy")
                dtpReturnDate.backColor = RGB(240, 255, 240)  ' Light green
            Else
                ' Highlight invalid date
                dtpReturnDate.backColor = RGB(255, 240, 240)  ' Light red
                MsgBox "Invalid date format. Use dd.mm.yyyy.", vbExclamation
                Cancel = True  ' Prevent leaving the field
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub


Private Sub ShowSearchHints()
    ' NEW FUNCTION: Search hints for the user
    Dim hintsText As String
    hintsText = "SEARCH HINTS:" & vbCrLf & vbCrLf
    hintsText = hintsText & "• To search for a sum, enter only numbers: 125000" & vbCrLf
    hintsText = hintsText & "• The system will find '125000', '125 000', '125000 rub.'" & vbCrLf
    hintsText = hintsText & "• Search works across all columns simultaneously" & vbCrLf
    hintsText = hintsText & "• You can search by part of a word or number" & vbCrLf
    hintsText = hintsText & vbCrLf & "Click 'Refresh data' if you modified Excel manually"
    
    MsgBox hintsText, vbInformation, "Search Help"
End Sub


'=====================================================================
'      MISSING STYLING PROCEDURES for frmLetterHistory
'=====================================================================
'=====================================================================
'      MISSING PROCEDURES in frmLetterHistory
'=====================================================================

Private Sub StyleButtonSafe(buttonName As String, buttonCaption As String, backColor As Long)
    ' Safe button styling for frmLetterHistory
    On Error Resume Next
    
    Dim btn As Object
    Set btn = Me.Controls(buttonName)
    
    If Not btn Is Nothing Then
        With btn
            .Caption = buttonCaption
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .backColor = backColor
            .ForeColor = RGB(255, 255, 255)
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub StyleLabelSafe(labelName As String)
    ' Safe label styling for frmLetterHistory
    On Error Resume Next
    
    Dim lbl As Object
    Set lbl = Me.Controls(labelName)
    
    If Not lbl Is Nothing Then
        With lbl
            .Font.Name = "Segoe UI"
            .Font.Size = 10
        End With
    End If
    
    On Error GoTo 0
End Sub

' FIXED ApplyElementStyles for frmLetterHistory
Private Sub ApplyElementStyles()
    On Error Resume Next
    
    If Not txtHistorySearch Is Nothing Then
        With txtHistorySearch
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .ControlTipText = "Search by addressee, number, date, attachments, executor" & vbCrLf & _
                              "To search by sum, enter numbers only (e.g.: 125000)"
        End With
    End If
    
    If Not lstLetterHistory Is Nothing Then
        With lstLetterHistory
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .backColor = RGB(255, 255, 255)
            .BorderStyle = 1
            .ControlTipText = "Double click on a letter to jump to the table record"
        End With
    End If
    
    If Not chkReceived Is Nothing Then
        With chkReceived
            .Caption = "Document received back"
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .ForeColor = RGB(0, 120, 0)
        End With
    End If
    
    ' Add to ApplyElementStyles for frmLetterHistory:
    StyleButtonSafe "btnSearchHelp", "Search help", RGB(158, 158, 158)

    
    ' NOW CORRECT: Button styling via local procedures
    StyleButtonSafe "btnUpdateStatus", "Update status", RGB(76, 175, 80)
    StyleButtonSafe "btnRefresh", "Refresh data", RGB(33, 150, 243)
    StyleButtonSafe "btnClose", "Close", RGB(244, 67, 54)
    StyleButtonSafe "btnClearSearch", "Clear search", RGB(255, 152, 0)
    StyleButtonSafe "btnExportToExcel", "Export to Excel", RGB(255, 152, 0)
    StyleButtonSafe "btnNavigateToRecord", "Go to record", RGB(103, 58, 183)
    
    ' Label styling via local procedures
    StyleLabelSafe "lblSearchInfo"
    StyleLabelSafe "lblSumLabel"
    StyleLabelSafe "lblDateLabel"
    
    On Error GoTo 0
End Sub


Private Sub SetControlTip(controlName As String, tipText As String)
    ' Setting tooltips for controls
    On Error Resume Next
    
    Dim ctrl As Object
    Set ctrl = Me.Controls(controlName)
    
    If Not ctrl Is Nothing Then
        ctrl.ControlTipText = tipText
        Debug.Print "Tooltip set for: " & controlName
    End If
    
    On Error GoTo 0
End Sub

