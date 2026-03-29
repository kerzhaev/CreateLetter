VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLetterHistory 
   Caption         =   "Letter History v1.3.1"
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
' Form: frmLetterHistory v1.3.1 - Thin-shell history UI with typed history records
' Author: CreateLetter contributors
' Date: 29.03.2026
' Purpose: History of sent letters with typed DTO bindings, thin-shell UI, and schema-safe status updates
' Updates v1.3.1:
' - Fixed VBA ByRef index mismatch for typed history record lookup helper
' - Preserved typed DTO contract and thin-shell history bindings
' Updates v1.3.0:
' - Switched history data flow from pipe-delimited strings to clsLetterHistoryRecord objects
' - Kept UI binding, search, export, and navigation on shared ModuleMain facade contracts
' - Preserved localized captions, Russian date formatting, and workbook-safe navigation
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
    ApplyLocalizedStaticCaptions
    ApplyElementStyles
    ConfigureDateFieldRussianFormat  ' NEW: Russian/European date format
    ConfigureCompactSumField         ' NEW: Compact sum field
    
    LoadAllLettersData
    ShowAllLettersOnInit
    InitializeControlValues
    
        Debug.Print "Letter history form initialized v1.3.1 with typed history records"
End Sub

Private Sub ConfigureDateFieldRussianFormat()
    ' FIXED FUNCTION: Date field configuration without non-existent properties
    On Error Resume Next
    
    If Not dtpReturnDate Is Nothing Then
        With dtpReturnDate
            ' FIXED: Setting the date in Russian/European format via Format()
            .value = Format(Date, "dd.mm.yyyy")
            
            ' REMOVED: .Format = 2 (not supported by TextBox)
            ' REMOVED: .CustomFormat = "dd.mm.yyyy" (not supported by TextBox)
            
            .ControlTipText = t("form.letter_history.tip.return_date", "Document return date (dd.mm.yyyy)")
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
            .ControlTipText = t("form.letter_history.tip.document_sum", "Document sum in rubles (numbers only or brief comment)")
        End With
        Debug.Print "Document sum field configured in compact mode"
    End If
    
    On Error GoTo 0
End Sub


Private Sub ApplyFormSettings()
    With Me
        .Caption = t("form.letter_history.title", "Letter History") & " v1.3.1"
        .backColor = RGB(250, 250, 250)
    End With
End Sub

Private Sub ApplyLocalizedStaticCaptions()
    On Error Resume Next

    SetLocalizedHistoryCaption "frameSearch", "form.letter_history.frame.search", "Search"
    SetLocalizedHistoryCaption "frameLettersList", "form.letter_history.frame.letters_list", "Letter list"
    SetLocalizedHistoryCaption "frameStatusEdit", "form.letter_history.frame.status_update", "Status update"
    SetLocalizedHistoryCaption "frameActions", "form.letter_history.frame.actions", "Actions"
    SetLocalizedHistoryCaption "Label1", "form.letter_history.label.search_letters", "Search letters by delivery history"
    SetLocalizedHistoryCaption "lblSearchLabel", "form.letter_history.label.search_status", "Search status"
    SetLocalizedHistoryCaption "lblSearchInfo", "status.ready", "Ready"
    SetLocalizedHistoryCaption "lblDateLabel", "form.letter_history.label.return_date", "Return date"
    SetLocalizedHistoryCaption "lblSumLabel", "form.letter_history.label.amount", "Amount"
    SetLocalizedHistoryCaption "btnUpdateStatus", "form.letter_history.caption.update_status", "Update status"
    SetLocalizedHistoryCaption "btnClose", "form.letter_history.caption.close", "Close"
    SetLocalizedHistoryCaption "btnRefresh", "form.letter_history.caption.refresh_data", "Refresh data"
    SetLocalizedHistoryCaption "btnClearSearch", "form.letter_history.caption.clear_search", "Clear search"
    SetLocalizedHistoryCaption "btnExportToExcel", "form.letter_history.caption.export_to_excel", "Export to Excel"
    SetLocalizedHistoryCaption "btnNavigateToRecord", "form.letter_history.caption.go_to_record", "Go to record"
    SetLocalizedHistoryCaption "btnSearchHelp", "form.letter_history.caption.search_help", "Search help"
    SetLocalizedHistoryCaption "chkReceived", "form.letter_history.caption.received_back", "Received back"

    On Error GoTo 0
End Sub

Private Sub SetLocalizedHistoryCaption(controlName As String, localizationKey As String, fallbackText As String)
    SetHistoryControlCaption controlName, t(localizationKey, fallbackText)
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
    
    Me.Controls("txtHistorySearch").value = ""
    Me.Controls("txtSumDocument").value = ""
    Me.Controls("chkReceived").value = False
    
    ' FIXED: Formatting date when setting value
    Me.Controls("dtpReturnDate").value = Format(Date, "dd.mm.yyyy")
    
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
        UpdateSearchInfo t("form.letter_history.msg.no_data", "No data found in worksheet 'Letters'")
        Exit Sub
    End If
    
    UpdateSearchInfo BuildHistoryLoadedCaption(allLettersData.count)
End Sub



Private Sub ShowAllLettersOnInit()
    If lstLetterHistory Is Nothing Then Exit Sub
    
    Set filteredData = FilterLetterHistoryRecords(allLettersData, "")
    BindHistoryList filteredData
    
    UpdateSearchInfo BuildHistoryShowingAllCaption(allLettersData.count)
End Sub

' ===============================================================================
' SEARCH AND FILTERING (NO CHANGES)
' ===============================================================================
Private Sub txtHistorySearch_Change()
    If txtHistorySearch Is Nothing Then Exit Sub
    
    Dim searchText As String
    searchText = Trim(txtHistorySearch.value)
    
    ' DEBUG: Outputting search information
    If IsNumeric(searchText) And Len(searchText) > 2 Then
        Debug.Print "=== SEARCH DEBUG ==="
        Debug.Print "Searching for number: " & searchText
        Debug.Print "Total records to search: " & allLettersData.count
        
        ' Show first few records for testing
        Dim i As Integer
        For i = 1 To WorksheetFunction.Min(3, allLettersData.count)
            Dim debugRecord As clsLetterHistoryRecord
            Set debugRecord = GetHistoryRecordFromCollection(allLettersData, i)
            If Not debugRecord Is Nothing Then
                Debug.Print "Record " & i & ", sum column: '" & debugRecord.DocumentSum & "'"
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
            UpdateSearchInfo BuildHistoryAmountSearchCaption(searchText)
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
        Dim letterRecord As clsLetterHistoryRecord
        Set letterRecord = GetHistoryRecordFromCollection(filteredData, selectedIndex)
        If Not letterRecord Is Nothing Then
            On Error Resume Next

            If Not txtSumDocument Is Nothing Then
                txtSumDocument.value = letterRecord.DocumentSum
            End If

            ParseReturnStatus letterRecord.ReturnStatus

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
        MsgBox t("form.letter_history.msg.select_record", "Select a letter to navigate to the record."), vbExclamation, t("form.letter_history.caption.go_to_record", "Go to record")
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.ListIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterRecord As clsLetterHistoryRecord
        Set letterRecord = GetHistoryRecordFromCollection(filteredData, selectedIndex)
        If Not letterRecord Is Nothing Then
            Dim rowNumber As Long
            rowNumber = letterRecord.RowNumber
            
            ' Getting "Letters" sheet
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets("Letters")
            
            If ws Is Nothing Then
                MsgBox t("form.letter_history.msg.letters_sheet_missing", "Worksheet 'Letters' not found."), vbCritical, t("form.letter_history.msg.navigation_error_title", "Navigation error")
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
            Application.StatusBar = t("form.letter_history.msg.selected_record", "Selected record: ") & letterRecord.Addressee & " | " & letterRecord.OutgoingNumber & " | " & letterRecord.OutgoingDate
            
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
    MsgBox t("form.letter_history.msg.navigation_error", "Error navigating to record: ") & Err.description, vbCritical, t("form.letter_history.msg.navigation_error_title", "Navigation error")
End Sub


Private Sub ParseReturnStatus(returnStatus As String)
    On Error Resume Next
    
    If HasReturnStatusDate(returnStatus) Then
        If Not chkReceived Is Nothing Then
            chkReceived.value = True
        End If
        
        Dim dateString As String
        dateString = ExtractReturnStatusDate(returnStatus)
        
        If IsDate(dateString) And Not dtpReturnDate Is Nothing Then
            ' FIXED: Formatting date in Russian/European format
            dtpReturnDate.value = Format(CDate(dateString), "dd.mm.yyyy")
        ElseIf Not dtpReturnDate Is Nothing Then
            dtpReturnDate.value = Format(Date, "dd.mm.yyyy")
        End If
    Else
        If Not chkReceived Is Nothing Then
            chkReceived.value = False
        End If
        If Not dtpReturnDate Is Nothing Then
            ' FIXED: Formatting date in Russian/European format
            dtpReturnDate.value = Format(Date, "dd.mm.yyyy")
        End If
    End If
    
    On Error GoTo 0
End Sub

' ===============================================================================
' ACTION BUTTONS (NO CHANGES)
' ===============================================================================
Private Sub btnClearSearch_Click()
    On Error Resume Next
    
    Dim originalCaption As String
    originalCaption = Me.Controls("btnClearSearch").Caption
    Me.Controls("btnClearSearch").Caption = t("form.letter_history.caption.clearing", "Clearing...")
    Me.Controls("btnClearSearch").Enabled = False
    
    DoEvents
    
    Me.Controls("txtHistorySearch").value = ""
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
    
    Me.Controls("txtSumDocument").value = ""
    Me.Controls("chkReceived").value = False
    
    ' FIXED: Format date upon clearing
    Me.Controls("dtpReturnDate").value = Format(Date, "dd.mm.yyyy")
    
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
        MsgBox t("form.letter_history.msg.select_status_update", "Select a letter to update the status."), vbExclamation
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.ListIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterRecord As clsLetterHistoryRecord
        Set letterRecord = GetHistoryRecordFromCollection(filteredData, selectedIndex)
        If Not letterRecord Is Nothing Then
            Dim rowNumber As Long
            rowNumber = letterRecord.RowNumber
            Dim returnStatus As String
            returnStatus = BuildLetterReturnStatus((Not chkReceived Is Nothing And chkReceived.value), ControlValueOrDefault("dtpReturnDate"))
            
            UpdateLetterHistoryRow rowNumber, ControlValueOrDefault("txtSumDocument"), returnStatus
            
            LoadAllLettersData
            txtHistorySearch_Change
            
            MsgBox t("form.letter_history.msg.status_updated", "Letter status updated successfully."), vbInformation
        End If
    End If
End Sub

Private Sub BindHistoryList(records As Collection)
    If lstLetterHistory Is Nothing Then Exit Sub
    
    lstLetterHistory.Clear
    
    Dim i As Long
    For i = 1 To records.count
        lstLetterHistory.AddItem FormatLetterHistoryDisplay(records(i))
    Next i
    
    UpdateSearchInfo BuildHistoryFoundCaption(records.count, allLettersData.count)
End Sub

Private Function ControlValueOrDefault(controlName As String, Optional defaultValue As String = "") As String
    On Error Resume Next
    ControlValueOrDefault = Trim(CStr(Me.Controls(controlName).value))
    If Err.number <> 0 Then
        ControlValueOrDefault = defaultValue
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub btnRefresh_Click()
    LoadAllLettersData
    txtHistorySearch_Change
    MsgBox t("form.letter_history.msg.data_refreshed", "Data refreshed."), vbInformation
End Sub

Private Sub btnExportToExcel_Click()
    ExportLetterHistoryRecords filteredData
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
        inputText = Trim(dtpReturnDate.value)
        
        ' If field is not empty, validate date
        If Len(inputText) > 0 Then
            If IsDate(inputText) Then
                ' Format valid date in Russian/European format
                dtpReturnDate.value = Format(CDate(inputText), "dd.mm.yyyy")
                dtpReturnDate.backColor = RGB(240, 255, 240)  ' Light green
            Else
                ' Highlight invalid date
                dtpReturnDate.backColor = RGB(255, 240, 240)  ' Light red
                MsgBox t("form.letter_history.msg.invalid_date", "Invalid date format. Use dd.mm.yyyy."), vbExclamation
                Cancel = True  ' Prevent leaving the field
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub


Private Sub ShowSearchHints()
    MsgBox GetLetterHistorySearchHintsText(), vbInformation, t("form.letter_history.msg.search_hints_title", "Search Help")
End Sub

Private Function GetHistoryRecordFromCollection(records As Collection, ByVal oneBasedIndex As Long) As clsLetterHistoryRecord
    On Error GoTo LookupFailed

    If records Is Nothing Then Exit Function
    If oneBasedIndex < 1 Or oneBasedIndex > records.Count Then Exit Function

    If IsObject(records(oneBasedIndex)) Then
        If TypeName(records(oneBasedIndex)) = "clsLetterHistoryRecord" Then
            Set GetHistoryRecordFromCollection = records(oneBasedIndex)
        End If
    End If

    Exit Function

LookupFailed:
    Set GetHistoryRecordFromCollection = Nothing
End Function


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
            .ControlTipText = t("form.letter_history.tip.search", "Search by addressee, number, date, attachments, executor" & vbCrLf & "To search by sum, enter numbers only (e.g.: 125000)")
        End With
    End If
    
    If Not lstLetterHistory Is Nothing Then
        With lstLetterHistory
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .backColor = RGB(255, 255, 255)
            .BorderStyle = 1
            .ControlTipText = t("form.letter_history.tip.double_click", "Double click on a letter to jump to the table record")
        End With
    End If
    
    If Not chkReceived Is Nothing Then
        With chkReceived
            .Caption = t("form.letter_history.caption.received_back", "Document received back")
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .ForeColor = RGB(0, 120, 0)
        End With
    End If
    
    ' Add to ApplyElementStyles for frmLetterHistory:
    StyleButtonSafe "btnSearchHelp", t("form.letter_history.caption.search_help", "Search help"), RGB(158, 158, 158)

    
    ' NOW CORRECT: Button styling via local procedures
    StyleButtonSafe "btnUpdateStatus", t("form.letter_history.caption.update_status", "Update status"), RGB(76, 175, 80)
    StyleButtonSafe "btnRefresh", t("form.letter_history.caption.refresh_data", "Refresh data"), RGB(33, 150, 243)
    StyleButtonSafe "btnClose", t("form.letter_history.caption.close", "Close"), RGB(244, 67, 54)
    StyleButtonSafe "btnClearSearch", t("form.letter_history.caption.clear_search", "Clear search"), RGB(255, 152, 0)
    StyleButtonSafe "btnExportToExcel", t("form.letter_history.caption.export_to_excel", "Export to Excel"), RGB(255, 152, 0)
    StyleButtonSafe "btnNavigateToRecord", t("form.letter_history.caption.go_to_record", "Go to record"), RGB(103, 58, 183)
    
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

