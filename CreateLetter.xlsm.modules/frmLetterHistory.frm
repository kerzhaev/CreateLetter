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
' Form: frmLetterHistory v1.2.0 - REVISED VERSION
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Date: 14.09.2025
' Purpose: History of sent letters with navigation to records
' Updates v1.2.0:
' - Reduced document sum field to a compact size
' - Added navigation to records in the "Letters" sheet on click
' - Set Russian/European date format for dtpReturnDate
' - Improved integration with the Excel table
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
    
    Debug.Print "Letter history form initialized v1.2.0 with improvements"
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
        .Caption = "Letter History v1.2.0"
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
    Set allLettersData = New Collection
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Letters")
    On Error GoTo 0
    
    If ws Is Nothing Then
        UpdateSearchInfo "Worksheet 'Letters' not found"
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        UpdateSearchInfo "No data found in worksheet 'Letters'"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To lastRow
        Dim letterData As String
        letterData = GetCellValueSafe(ws, i, 1) & "|" & _
                     GetCellValueSafe(ws, i, 2) & "|" & _
                     GetCellValueSafe(ws, i, 3) & "|" & _
                     GetCellValueSafe(ws, i, 4) & "|" & _
                     GetCellValueSafe(ws, i, 5) & "|" & _
                     GetCellValueSafe(ws, i, 6) & "|" & _
                     GetCellValueSafe(ws, i, 7) & "|" & _
                     GetCellValueSafe(ws, i, 8) & "|" & _
                     CStr(i)
        
        allLettersData.Add letterData
    Next i
    
    UpdateSearchInfo "Letters loaded: " & allLettersData.count
End Sub



Private Sub ShowAllLettersOnInit()
    If lstLetterHistory Is Nothing Then Exit Sub
    
    Set filteredData = New Collection
    lstLetterHistory.Clear
    
    Dim i As Integer
    For i = 1 To allLettersData.count
        Dim letterData As String
        letterData = allLettersData(i)
        
        filteredData.Add letterData
        Dim displayText As String
        displayText = FormatLetterForDisplay(letterData)
        lstLetterHistory.AddItem displayText
    Next i
    
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
                Debug.Print "Record " & i & ", sum column: '" & parts(4) & "'"
            End If
        Next i
        Debug.Print "=================="
    End If
    
    If lstLetterHistory Is Nothing Then Exit Sub
    lstLetterHistory.Clear
    
    If Len(searchText) = 0 Then
        ShowAllLettersOnInit
    Else
        DisplayFilteredLetters searchText
        
        If IsNumeric(searchText) And Len(searchText) > 2 Then
            UpdateSearchInfo "Searching for number " & searchText & " in document amounts..."
        End If
    End If
End Sub



Private Sub DisplayFilteredLetters(searchText As String)
    Set filteredData = New Collection
    If lstLetterHistory Is Nothing Then Exit Sub
    lstLetterHistory.Clear
    
    Dim foundCount As Integer
    foundCount = 0
    
    Dim i As Integer
    For i = 1 To allLettersData.count
        Dim letterData As String
        letterData = allLettersData(i)
        
        Dim isMatch As Boolean
        isMatch = False
        
        Dim parts() As String
        parts = Split(letterData, "|")
        
        Dim j As Integer
        For j = 0 To UBound(parts) - 1
            Dim searchInText As String
            searchInText = UCase(parts(j))
            Dim searchPattern As String
            searchPattern = UCase(searchText)
            
            ' FIXED: Special handling for searching by sum (column 4)
            If j = 4 Then  ' Document sum column
                If IsNumeric(searchPattern) Then
                    ' FIXED: Improved search for numbers
                    If IsNumericMatch(searchInText, searchPattern) Then
                        isMatch = True
                        Exit For
                    End If
                Else
                    ' If searching for non-number, search as usual
                    If InStr(searchInText, searchPattern) > 0 Then
                        isMatch = True
                        Exit For
                    End If
                End If
            Else
                ' Standard search across other columns
                If InStr(searchInText, searchPattern) > 0 Then
                    isMatch = True
                    Exit For
                End If
            End If
        Next j
        
        If isMatch Then
            filteredData.Add letterData
            Dim displayText As String
            displayText = FormatLetterForDisplay(letterData)
            lstLetterHistory.AddItem displayText
            foundCount = foundCount + 1
        End If
    Next i
    
    UpdateSearchInfo "Letters found: " & foundCount & " of " & allLettersData.count
End Sub

' NEW FUNCTION: Improved number comparison
Private Function IsNumericMatch(cellValue As String, searchValue As String) As Boolean
    IsNumericMatch = False
    
    ' FIXED: More aggressive cleaning of all non-numeric characters
    Dim cleanCellValue As String
    cleanCellValue = ExtractOnlyDigits(cellValue)
    
    Dim cleanSearchValue As String
    cleanSearchValue = ExtractOnlyDigits(searchValue)
    
    ' If both values are empty after cleaning - no match
    If Len(cleanCellValue) = 0 Or Len(cleanSearchValue) = 0 Then Exit Function
    
    ' 1. Exact digit match
    If cleanCellValue = cleanSearchValue Then
        IsNumericMatch = True
        Exit Function
    End If
    
    ' 2. Partial match (substring search within numbers)
    If InStr(cleanCellValue, cleanSearchValue) > 0 Then
        IsNumericMatch = True
        Exit Function
    End If
    
    Debug.Print "Comparison: '" & cleanCellValue & "' vs '" & cleanSearchValue & "'"
End Function

' NEW FUNCTION: Extract ONLY digits
Private Function ExtractOnlyDigits(inputText As String) As String
    ExtractOnlyDigits = ""
    
    If Len(inputText) = 0 Then Exit Function
    
    Dim i As Integer
    For i = 1 To Len(inputText)
        Dim char As String
        char = Mid(inputText, i, 1)
        
        ' Take ONLY digits, ignore everything else
        If char >= "0" And char <= "9" Then
            ExtractOnlyDigits = ExtractOnlyDigits & char
        End If
    Next i
    
    Debug.Print "Extracted from '" & inputText & "': '" & ExtractOnlyDigits & "'"
End Function



' FIXED FUNCTION: More reliable digit extraction
Private Function ExtractNumbersOnly(inputText As String) As String
    ExtractNumbersOnly = ""
    
    If Len(inputText) = 0 Then Exit Function
    
    Dim i As Integer
    For i = 1 To Len(inputText)
        Dim char As String
        char = Mid(inputText, i, 1)
        
        ' FIXED: Accounting for all types of spaces and separators
        If char >= "0" And char <= "9" Then
            ExtractNumbersOnly = ExtractNumbersOnly & char
        ElseIf char = " " Or char = Chr(160) Or char = "," Or char = "." Then
            ' Ignore spaces (regular and non-breaking), commas, and dots
            ' Chr(160) is a non-breaking space
        End If
    Next i
End Function


' ADDITIONAL FUNCTION: Improved data loading with correct formatting
Private Function GetCellValueSafe(ws As Worksheet, Row As Long, col As Long) As String
    On Error Resume Next
    
    Dim cellValue As Variant
    cellValue = ws.Cells(Row, col).Value
    
    ' FIXED: Special handling for numeric values
    If col = 5 Then  ' Document sum column
        If IsNumeric(cellValue) And cellValue <> 0 Then
            ' Format number without decimals if it's a whole number
            If cellValue = Int(cellValue) Then
                GetCellValueSafe = CStr(CLng(cellValue))
            Else
                GetCellValueSafe = CStr(cellValue)
            End If
        Else
            GetCellValueSafe = CStr(cellValue)
        End If
    Else
        GetCellValueSafe = CStr(cellValue)
    End If
    
    If Err.number <> 0 Then
        GetCellValueSafe = ""
    End If
    
    On Error GoTo 0
End Function



Private Function FormatLetterForDisplay(letterData As String) As String
    Dim parts() As String
    parts = Split(letterData, "|")
    
    If UBound(parts) >= 8 Then
        Dim formattedDate As String
        formattedDate = FormatDisplayDate(parts(2))
        
        Dim formattedSum As String
        If Len(Trim(parts(4))) > 0 And IsNumeric(parts(4)) Then
            If CDbl(parts(4)) > 0 Then
                formattedSum = Format(CDbl(parts(4)), "#,##0.00") & " rub."
            Else
                formattedSum = "—"
            End If
        Else
            formattedSum = "—"
        End If
        
        Dim statusIcon As String
        If InStr(UCase(parts(5)), "RECEIVED") > 0 And InStr(UCase(parts(5)), "NOT RECEIVED") = 0 Then
            statusIcon = "? " & parts(5)
        Else
            statusIcon = "? " & parts(5)
        End If
        
        Dim addressee As String, attachments As String
        addressee = Left(parts(0), 25) & IIf(Len(parts(0)) > 25, "...", "")
        attachments = Left(parts(3), 30) & IIf(Len(parts(3)) > 30, "...", "")
        
        FormatLetterForDisplay = addressee & " | " & _
                                parts(1) & " | " & _
                                formattedDate & " | " & _
                                attachments & " | " & _
                                formattedSum & " | " & _
                                statusIcon & " | " & _
                                parts(6) & " | " & _
                                parts(7)
    Else
        FormatLetterForDisplay = letterData
    End If
End Function

Private Function FormatDisplayDate(dateValue As String) As String
    On Error Resume Next
    If IsDate(dateValue) Then
        FormatDisplayDate = Format(CDate(dateValue), "dd.mm.yyyy")
    Else
        FormatDisplayDate = dateValue
    End If
    On Error GoTo 0
End Function

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
        parts = Split(letterData, "|")
        
        If UBound(parts) >= 8 Then
            On Error Resume Next
            
            If Not txtSumDocument Is Nothing Then
                txtSumDocument.Value = parts(4)
            End If
            
            ParseReturnStatus parts(5)
            
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
        parts = Split(letterData, "|")
        
        If UBound(parts) >= 8 Then
            Dim rowNumber As Long
            rowNumber = CLng(parts(8))
            
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
            ws.Cells(rowNumber, 1).Select
            
            ' Highlight record row
            With ws.Rows(rowNumber).Interior
                .Color = RGB(255, 255, 0)  ' Yellow highlight
                .Pattern = xlSolid
            End With
            
            ' NEW: Show info in Excel status bar
            Application.StatusBar = "Selected record: " & parts(0) & " | " & parts(1) & " | " & parts(2)
            
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
        parts = Split(letterData, "|")
        
        If UBound(parts) >= 8 Then
            Dim rowNumber As Long
            rowNumber = CLng(parts(8))
            
            Dim ws As Worksheet
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets("Letters")
            On Error GoTo 0
            
            If ws Is Nothing Then
                MsgBox "Worksheet 'Letters' not found.", vbExclamation
                Exit Sub
            End If
            
            On Error Resume Next
            If Not txtSumDocument Is Nothing Then
                Dim sumValue As String
                sumValue = Trim(txtSumDocument.Value)
                
                If Len(sumValue) = 0 Then
                    ws.Cells(rowNumber, 5).Value = ""
                ElseIf IsNumeric(sumValue) Then
                    ws.Cells(rowNumber, 5).Value = CDbl(sumValue)
                Else
                    ws.Cells(rowNumber, 5).Value = sumValue
                End If
            End If
            
            ' Find this piece of code and fix it:
            Dim returnStatus As String
            If Not chkReceived Is Nothing And chkReceived.Value Then
                If Not dtpReturnDate Is Nothing Then
                    ' FIXED: Parsing date from text field
                    Dim dateValue As Date
                    If IsDate(dtpReturnDate.Value) Then
                        dateValue = CDate(dtpReturnDate.Value)
                    Else
                        dateValue = Date
                    End If
                    returnStatus = Format(dateValue, "dd.mm.yyyy") & " received"
                Else
                    returnStatus = Format(Date, "dd.mm.yyyy") & " received"
                End If
            Else
                returnStatus = "not received"
            End If

            
            ws.Cells(rowNumber, 6).Value = returnStatus
            On Error GoTo 0
            
            LoadAllLettersData
            txtHistorySearch_Change
            
            MsgBox "Letter status updated successfully.", vbInformation
        End If
    End If
End Sub

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
        .Cells(1, 1).Value = "Addressee"
        .Cells(1, 2).Value = "Outgoing Number"
        .Cells(1, 3).Value = "Outgoing Date"
        .Cells(1, 4).Value = "Attachment Name"
        .Cells(1, 5).Value = "Document Sum"
        .Cells(1, 6).Value = "Return Mark"
        .Cells(1, 7).Value = "Executor Name"
        .Cells(1, 8).Value = "Send Type"
        
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
        
        If UBound(parts) >= 7 Then
            exportWs.Cells(i + 1, 1).Value = parts(0)
            exportWs.Cells(i + 1, 2).Value = parts(1)
            exportWs.Cells(i + 1, 3).Value = parts(2)
            exportWs.Cells(i + 1, 4).Value = parts(3)
            exportWs.Cells(i + 1, 5).Value = parts(4)
            exportWs.Cells(i + 1, 6).Value = parts(5)
            exportWs.Cells(i + 1, 7).Value = parts(6)
            exportWs.Cells(i + 1, 8).Value = parts(7)
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

