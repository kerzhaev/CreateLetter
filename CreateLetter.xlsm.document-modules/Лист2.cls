VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ======================================================================
' SHEET MODULE: "Letters"
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Opening the letter history form by clicking on a letter row
' ======================================================================
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' Double-click event to open the letter history form
    On Error GoTo DoubleClickError
    
    Dim selectedRow As Long
    
    ' Check that this is the "Letters" sheet and only one cell is selected
    If Me.Name = "Letters" And Target.Cells.count = 1 Then
        selectedRow = Target.Row
        
        ' Skip the header row (first row)
        If selectedRow > 1 And Len(Trim(Me.Cells(selectedRow, 1).value)) > 0 Then
            Cancel = True  ' Cancel standard double-click behavior
            
            ' Gather data from the selected letter row
            Dim letterData(7) As String
            Dim i As Integer
            
            For i = 0 To 7
                letterData(i) = CStr(Me.Cells(selectedRow, i + 1).value)
            Next i
            
            ' Open the history form with pre-filled data
            OpenLetterHistoryWithData letterData, selectedRow
            
            ' Show status message
            SetStatusBarMessage "History opened for letter: " & letterData(1) & " dated " & letterData(2), 3
            
            Debug.Print "Letter history opened: " & letterData(0) & " No. " & letterData(1) & " (row " & selectedRow & ")"
        End If
    End If
    
    Exit Sub
    
DoubleClickError:
    MsgBox "Error opening letter history form: " & Err.description, vbCritical, "Error"
End Sub

Private Sub OpenLetterHistoryWithData(letterArray As Variant, rowNumber As Long)
    ' Procedure to open the history form with pre-filled data
    On Error GoTo OpenHistoryError
    
    ' Check if the form is already open
    Dim existingForm As Object
    On Error Resume Next
    Set existingForm = VBA.UserForms("frmLetterHistory")
    On Error GoTo OpenHistoryError
    
    If Not existingForm Is Nothing Then
        ' If the form is already open, fill it with new data
        FillHistoryFormWithData existingForm, letterArray, rowNumber
        existingForm.SetFocus
        existingForm.ZOrder 0
        Debug.Print "Existing history form updated with data from row " & rowNumber
    Else
        ' Load a new form
        Load frmLetterHistory
        FillHistoryFormWithData frmLetterHistory, letterArray, rowNumber
        frmLetterHistory.Show vbModeless
        Debug.Print "New history form created with data from row " & rowNumber
    End If
    
    Exit Sub
    
OpenHistoryError:
    MsgBox "Error opening history form: " & Err.description, vbCritical, "Error"
End Sub

Private Sub FillHistoryFormWithData(targetForm As Object, letterArray As Variant, rowNumber As Long)
    ' IMPROVED VERSION: Filling with full type safety
    On Error Resume Next
    
    With targetForm
        ' 1. AUTO-SEARCH: Fill the search field with the addressee
        If Not .txtHistorySearch Is Nothing Then
            .txtHistorySearch.value = SafeGetStringValue(letterArray, 0)  ' Addressee
        End If
        
        ' 2. EDIT DATA: Fill status and sum fields
        If Not .txtSumDocument Is Nothing Then
            .txtSumDocument.value = SafeGetStringValue(letterArray, 4)    ' Document sum
        End If
        
        ' 3. RECEIVED STATUS: Check the return mark
        If Not .chkReceived Is Nothing Then
            Dim returnStatus As String
            returnStatus = UCase(SafeGetStringValue(letterArray, 5))
            .chkReceived.value = (InStr(returnStatus, "RECEIVED") > 0 And InStr(returnStatus, "NOT RECEIVED") = 0)
        End If
        
        ' 4. RETURN DATE: SAFE extraction from the return status
        If Not .dtpReturnDate Is Nothing Then
            Dim statusText As String
            Dim dateString As String
            
            statusText = SafeGetStringValue(letterArray, 5)
            dateString = ExtractDateFromStatusSafe(statusText)
            
            If IsDate(dateString) Then
                .dtpReturnDate.value = Format(CDate(dateString), "dd.mm.yyyy")
            Else
                .dtpReturnDate.value = Format(Date, "dd.mm.yyyy")
            End If
        End If
        
        ' 5. LETTER INFO: Save for future use
        .Tag = "Row:" & rowNumber & "|Number:" & SafeGetStringValue(letterArray, 1) & "|Date:" & SafeGetStringValue(letterArray, 2)
        
        ' 6. RUN SEARCH: Activate search by addressee
        .txtHistorySearch_Change
        
        Debug.Print "History form filled: Addressee=" & SafeGetStringValue(letterArray, 0) & ", Sum=" & SafeGetStringValue(letterArray, 4)
    End With
    
    On Error GoTo 0
End Sub

' NEW HELPER FUNCTION: Safe string value retrieval
Private Function SafeGetStringValue(arr As Variant, index As Integer) As String
    ' Safe extraction of a string value from an array
    SafeGetStringValue = ""
    
    On Error Resume Next
    
    If IsArray(arr) Then
        If index >= LBound(arr) And index <= UBound(arr) Then
            SafeGetStringValue = CStr(arr(index))
        End If
    End If
    
    On Error GoTo 0
End Function

' REWRITTEN FUNCTION: Safe date extraction
Private Function ExtractDateFromStatusSafe(ByVal statusText As String) As String
    ' SAFE version of date extraction from status
    ExtractDateFromStatusSafe = ""
    
    On Error Resume Next
    
    ' Check input data
    If Len(Trim(statusText)) = 0 Then Exit Function
    
    ' Look for the date in various formats
    Dim parts() As String
    parts = Split(statusText, " ")
    
    Dim i As Integer
    For i = 0 To UBound(parts)
        Dim testString As String
        testString = Trim(parts(i))
        
        ' Check different date formats
        If Len(testString) > 0 Then
            ' Format dd.mm.yyyy or dd/mm/yyyy
            If IsDate(testString) Then
                ExtractDateFromStatusSafe = testString
                Exit Function
            End If
            
            ' Format ddmmyyyy (no separators)
            If Len(testString) = 8 And IsNumeric(testString) Then
                Dim formattedDate As String
                formattedDate = Left(testString, 2) & "." & Mid(testString, 3, 2) & "." & Right(testString, 4)
                If IsDate(formattedDate) Then
                    ExtractDateFromStatusSafe = formattedDate
                    Exit Function
                End If
            End If
        End If
    Next i
    
    On Error GoTo 0
End Function


Private Function ExtractDateFromStatus(statusText As String) As String
    ' Extract date from return status text
    ExtractDateFromStatus = ""
    
    If Len(statusText) = 0 Then Exit Function
    
    ' Look for date in dd.mm.yyyy format
    Dim parts() As String
    parts = Split(statusText, " ")
    
    Dim i As Integer
    For i = 0 To UBound(parts)
        If IsDate(parts(i)) Then
            ExtractDateFromStatus = parts(i)
            Exit Function
        End If
    Next i
End Function