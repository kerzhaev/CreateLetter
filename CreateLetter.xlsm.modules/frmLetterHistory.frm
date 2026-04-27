VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLetterHistory 
   Caption         =   "История писем v1.3.5"
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
' Form: frmLetterHistory v1.3.5 - Thin-shell history UI with typed history records
' Author: CreateLetter contributors
' Date: 29.03.2026
' Purpose: History of sent letters with typed DTO bindings, thin-shell UI, and schema-safe status updates
' Updates v1.3.3:
' - Removed high-value Resume Next usage from record selection and return-status parsing paths
' - Kept styling-only helpers lightweight while making data-binding failures explicit
' Updates v1.3.2:
' - Replaced high-value history reset and date-validation Resume Next paths with targeted handlers
' - Kept styling-only helpers lightweight while making status/search flow failures explicit
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
    
        Debug.Print "Letter history form initialized v1.3.3 with typed history records"
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
            
            .ControlTipText = t("form.letter_history.tip.return_date", "Дата возврата документа (дд.мм.гггг)")
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
            .MultiLine = False              ' Single-line field
            .WordWrap = False               ' No word wrap
            .ScrollBars = 0                 ' No scrollbars (0 = fmScrollBarsNone)
            ' REMOVED: .EnterKeyBehaviour = False (not supported)
            .Height = 24                    ' Fixed height
            .ControlTipText = t("form.letter_history.tip.document_sum", "Сумма документа в рублях (только цифры или короткий комментарий)")
        End With
        Debug.Print "Document sum field configured in compact mode"
    End If
    
    On Error GoTo 0
End Sub


Private Sub ApplyFormSettings()
    With Me
        .Caption = t("form.letter_history.title", "История отправленных писем") & " v1.3.5"
        .backColor = RGB(250, 250, 250)
    End With
End Sub

Private Sub ApplyLocalizedStaticCaptions()
    On Error Resume Next

    SetLocalizedHistoryCaption "frameSearch", "form.letter_history.frame.search", "Поиск"
    SetLocalizedHistoryCaption "frameLettersList", "form.letter_history.frame.letters_list", "Список писем"
    SetLocalizedHistoryCaption "frameStatusEdit", "form.letter_history.frame.status_update", "Обновление статуса"
    SetLocalizedHistoryCaption "frameActions", "form.letter_history.frame.actions", "Действия"
    SetLocalizedHistoryCaption "Label1", "form.letter_history.label.search_letters", "Поиск писем по истории доставки"
    SetLocalizedHistoryCaption "lblSearchLabel", "form.letter_history.label.search_status", "Статус поиска"
    SetLocalizedHistoryCaption "lblSearchInfo", "status.ready", "Готово"
    SetLocalizedHistoryCaption "lblDateLabel", "form.letter_history.label.return_date", "Дата возврата"
    SetLocalizedHistoryCaption "lblSumLabel", "form.letter_history.label.amount", "Сумма"
    SetLocalizedHistoryCaption "btnUpdateStatus", "form.letter_history.caption.update_status", "Обновить статус"
    SetLocalizedHistoryCaption "btnClose", "form.letter_history.caption.close", "Закрыть"
    SetLocalizedHistoryCaption "btnRefresh", "form.letter_history.caption.refresh_data", "Обновить данные"
    SetLocalizedHistoryCaption "btnClearSearch", "form.letter_history.caption.clear_search", "Очистить поиск"
    SetLocalizedHistoryCaption "btnExportToExcel", "form.letter_history.caption.export_to_excel", "Экспорт в Excel"
    SetLocalizedHistoryCaption "btnNavigateToRecord", "form.letter_history.caption.go_to_record", "Перейти к записи"
    SetLocalizedHistoryCaption "btnSearchHelp", "form.letter_history.caption.search_help", "Справка по поиску"
    SetLocalizedHistoryCaption "chkReceived", "form.letter_history.caption.received_back", "Получено обратно"

    On Error GoTo 0
End Sub

Private Sub SetLocalizedHistoryCaption(controlName As String, localizationKey As String, fallbackText As String)
    SetHistoryControlCaption controlName, t(localizationKey, fallbackText)
End Sub

Private Sub SetHistoryControlCaption(controlName As String, captionText As String)
    On Error Resume Next

    Dim ctrl As control
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
        UpdateSearchInfo t("form.letter_history.msg.no_data", "На листе 'Letters' данные не найдены")
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
    On Error GoTo SelectionError

    If lstLetterHistory Is Nothing Then Exit Sub
    If lstLetterHistory.listIndex < 0 Then Exit Sub
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.listIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterRecord As clsLetterHistoryRecord
        Set letterRecord = GetHistoryRecordFromCollection(filteredData, selectedIndex)
        If Not letterRecord Is Nothing Then
            If Not txtSumDocument Is Nothing Then
                txtSumDocument.value = letterRecord.DocumentSum
            End If

            ParseReturnStatus letterRecord.returnStatus
        End If
    End If
    Exit Sub

SelectionError:
    MsgBox t("form.letter_history.msg.selection_error", "Ошибка при загрузке выбранной записи истории: ") & Err.description, vbExclamation
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
    If lstLetterHistory.listIndex < 0 Then
        MsgBox t("form.letter_history.msg.select_record", "Выберите письмо для перехода к записи."), vbExclamation, t("form.letter_history.caption.go_to_record", "Перейти к записи")
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.listIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterRecord As clsLetterHistoryRecord
        Set letterRecord = GetHistoryRecordFromCollection(filteredData, selectedIndex)
        If Not letterRecord Is Nothing Then
            Dim rowNumber As Long
            rowNumber = letterRecord.rowNumber
            
            ' Getting "Letters" sheet
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets("Letters")
            
            If ws Is Nothing Then
                MsgBox t("form.letter_history.msg.letters_sheet_missing", "Лист 'Letters' не найден."), vbCritical, t("form.letter_history.msg.navigation_error_title", "Ошибка навигации")
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
            Application.StatusBar = t("form.letter_history.msg.selected_record", "Выбрана запись: ") & letterRecord.Addressee & " | " & letterRecord.OutgoingNumber & " | " & letterRecord.OutgoingDate
            
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
    MsgBox t("form.letter_history.msg.navigation_error", "Ошибка при переходе к записи: ") & Err.description, vbCritical, t("form.letter_history.msg.navigation_error_title", "Ошибка навигации")
End Sub


Private Sub ParseReturnStatus(returnStatus As String)
    On Error GoTo ParseError

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
    Exit Sub

ParseError:
    If Not chkReceived Is Nothing Then
        chkReceived.value = False
    End If
    If Not dtpReturnDate Is Nothing Then
        dtpReturnDate.value = Format(Date, "dd.mm.yyyy")
    End If
    Debug.Print "ParseReturnStatus error: " & Err.description
End Sub

' ===============================================================================
' ACTION BUTTONS (NO CHANGES)
' ===============================================================================
Private Sub btnClearSearch_Click()
    On Error GoTo ClearError
    
    Dim originalCaption As String
    originalCaption = Me.Controls("btnClearSearch").Caption
    Me.Controls("btnClearSearch").Caption = t("form.letter_history.caption.clearing", "Очистка...")
    Me.Controls("btnClearSearch").Enabled = False
    
    DoEvents
    
    Me.Controls("txtHistorySearch").value = ""
    ClearAllHistoryFields
    ShowAllLettersOnInit
    
    Me.Controls("btnClearSearch").Caption = originalCaption
    Me.Controls("btnClearSearch").Enabled = True
    Me.Controls("txtHistorySearch").SetFocus
    
    Debug.Print "Full clear of letter history form executed"
    Exit Sub

ClearError:
    Me.Controls("btnClearSearch").Caption = t("form.letter_history.caption.clear_search", "Очистить поиск")
    Me.Controls("btnClearSearch").Enabled = True
    MsgBox t("form.letter_history.msg.clear_error", "Ошибка при очистке формы истории: ") & Err.description, vbExclamation
End Sub

Private Sub ClearAllHistoryFields()
    On Error GoTo ClearFieldsError
    
    Me.Controls("txtSumDocument").value = ""
    Me.Controls("chkReceived").value = False
    
    ' FIXED: Format date upon clearing
    Me.Controls("dtpReturnDate").value = Format(Date, "dd.mm.yyyy")
    
    SetControlBackColor "txtSumDocument", RGB(255, 255, 255)
    
    Debug.Print "All letter history fields cleared"
    Exit Sub

ClearFieldsError:
    Err.Raise Err.Number, "ClearAllHistoryFields", Err.description
End Sub


Private Sub SetControlBackColor(controlName As String, backColor As Long)
    On Error Resume Next
    
    Dim ctrl As control
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then
        ctrl.backColor = backColor
    End If
    
    On Error GoTo 0
End Sub

Private Sub btnUpdateStatus_Click()
    If lstLetterHistory Is Nothing Then Exit Sub
    If lstLetterHistory.listIndex < 0 Then
        MsgBox t("form.letter_history.msg.select_status_update", "Выберите письмо для обновления статуса."), vbExclamation
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstLetterHistory.listIndex + 1
    
    If selectedIndex <= filteredData.count Then
        Dim letterRecord As clsLetterHistoryRecord
        Set letterRecord = GetHistoryRecordFromCollection(filteredData, selectedIndex)
        If Not letterRecord Is Nothing Then
            Dim rowNumber As Long
            rowNumber = letterRecord.rowNumber
            Dim returnStatus As String
            returnStatus = BuildLetterReturnStatus((Not chkReceived Is Nothing And chkReceived.value), ControlValueOrDefault("dtpReturnDate"))
            
            UpdateLetterHistoryRow rowNumber, ControlValueOrDefault("txtSumDocument"), returnStatus
            
            LoadAllLettersData
            txtHistorySearch_Change
            
            MsgBox t("form.letter_history.msg.status_updated", "Статус письма успешно обновлен."), vbInformation
        End If
    End If
End Sub

Private Sub BindHistoryList(records As Collection)
    If lstLetterHistory Is Nothing Then Exit Sub
    
    lstLetterHistory.Clear
    lstLetterHistory.ColumnCount = 2
    lstLetterHistory.ColumnWidths = "530 pt;110 pt"
    
    Dim i As Long
    For i = 1 To records.count
        lstLetterHistory.AddItem FormatLetterHistoryDisplay(records(i))
        lstLetterHistory.List(lstLetterHistory.ListCount - 1, 1) = RepositoryGetLetterHistoryPackedStatusDisplay(records(i))
    Next i
    
    UpdateSearchInfo BuildHistoryFoundCaption(records.count, allLettersData.count)
End Sub

Private Function ControlValueOrDefault(controlName As String, Optional defaultValue As String = "") As String
    On Error GoTo ReadFailed
    ControlValueOrDefault = Trim(CStr(Me.Controls(controlName).value))
    Exit Function

ReadFailed:
    ControlValueOrDefault = defaultValue
End Function

Private Sub btnRefresh_Click()
    LoadAllLettersData
    txtHistorySearch_Change
    MsgBox t("form.letter_history.msg.data_refreshed", "Данные обновлены."), vbInformation
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
    On Error GoTo ValidationError
    
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
                MsgBox t("form.letter_history.msg.invalid_date", "Неверный формат даты. Используйте дд.мм.гггг."), vbExclamation
                Cancel = True  ' Prevent leaving the field
            End If
        End If
    End If
    Exit Sub

ValidationError:
    MsgBox t("form.letter_history.msg.invalid_date", "Неверный формат даты. Используйте дд.мм.гггг."), vbExclamation
    Cancel = True
End Sub


Private Sub ShowSearchHints()
    MsgBox GetLetterHistorySearchHintsText(), vbInformation, t("form.letter_history.msg.search_hints_title", "Справка по поиску")
End Sub

Private Function GetHistoryRecordFromCollection(records As Collection, ByVal oneBasedIndex As Long) As clsLetterHistoryRecord
    On Error GoTo LookupFailed

    If records Is Nothing Then Exit Function
    If oneBasedIndex < 1 Or oneBasedIndex > records.count Then Exit Function

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
            .ControlTipText = t("form.letter_history.tip.search", "Поиск по адресату, номеру, дате, приложениям, исполнителю" & vbCrLf & "Для поиска по сумме вводите только цифры (например: 125000)")
        End With
    End If
    
    If Not lstLetterHistory Is Nothing Then
        With lstLetterHistory
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .ColumnCount = 2
            .ColumnWidths = "530 pt;110 pt"
            .backColor = RGB(255, 255, 255)
            .BorderStyle = 1
            .ControlTipText = t("form.letter_history.tip.double_click", "Дважды щелкните по письму, чтобы перейти к записи в таблице")
        End With
    End If
    
    If Not chkReceived Is Nothing Then
        With chkReceived
            .Caption = t("form.letter_history.caption.received_back", "Документ получен обратно")
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .ForeColor = RGB(0, 120, 0)
        End With
    End If
    
    ' Add to ApplyElementStyles for frmLetterHistory:
    StyleButtonSafe "btnSearchHelp", t("form.letter_history.caption.search_help", "Справка по поиску"), RGB(158, 158, 158)

    
    ' NOW CORRECT: Button styling via local procedures
    StyleButtonSafe "btnUpdateStatus", t("form.letter_history.caption.update_status", "Обновить статус"), RGB(76, 175, 80)
    StyleButtonSafe "btnRefresh", t("form.letter_history.caption.refresh_data", "Обновить данные"), RGB(33, 150, 243)
    StyleButtonSafe "btnClose", t("form.letter_history.caption.close", "Закрыть"), RGB(244, 67, 54)
    StyleButtonSafe "btnClearSearch", t("form.letter_history.caption.clear_search", "Очистить поиск"), RGB(255, 152, 0)
    StyleButtonSafe "btnExportToExcel", t("form.letter_history.caption.export_to_excel", "Экспорт в Excel"), RGB(255, 152, 0)
    StyleButtonSafe "btnNavigateToRecord", t("form.letter_history.caption.go_to_record", "Перейти к записи"), RGB(103, 58, 183)
    
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

