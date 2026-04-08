VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ======================================================================
' SHEET MODULE: "Addresses"
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Opening the letter creation form by clicking on an address row
' ======================================================================
Option Explicit

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    ' Event triggered when the selected cell on the "Addresses" sheet changes
'    On Error GoTo SelectionError
'
'    Dim addressRow As Long
'
'    ' Check that only one cell is selected and it's our sheet
'    If Me.Name = "Addresses" And Target.Cells.Count = 1 Then
'        addressRow = Target.Row
'
'        ' Skip the header row (first row)
'        If addressRow > 1 Then
'
'            ' Check that the row contains addressee data
'            If Len(Trim(Me.Cells(addressRow, 1).Value)) > 0 Then
'
'                ' Gather data from the selected row
'                Dim addressData(7) As String
'                Dim i As Integer
'
'                For i = 0 To 7
'                    addressData(i) = CStr(Me.Cells(addressRow, i + 1).Value)
'                Next i
'
'                ' Open the form with pre-filled data
'                OpenLetterCreatorWithAddress addressData, addressRow
'
'                Debug.Print "Form opened for addressee: " & addressData(0) & " (row " & addressRow & ")"
'            End If
'        End If
'    End If
'
'
'     ' ADDITIONALLY: Highlighting the selected address row
'    If Me.Name = "Addresses" And Target.Cells.Count = 1 Then
'        addressRow = Target.Row
'
'        ' Clear previous highlight
'        Me.Cells.Interior.Pattern = xlNone
'
'        If addressRow > 1 And Len(Trim(Me.Cells(addressRow, 1).Value)) > 0 Then
'            ' Highlight the active row
'            Me.Rows(addressRow).Interior.Color = RGB(255, 255, 200)  ' Light yellow
'
'            ' Existing form opening code...
'            OpenLetterCreatorWithAddress addressData, addressRow
'
'            ' Remove highlight after 2 seconds
'            Application.OnTime Now + TimeValue("00:00:02"), "ClearAddressHighlight"
'        End If
'    End If
'
'
'
'
'    Exit Sub
'
'SelectionError:
'    Debug.Print "Error opening form: " & Err.description
'End Sub

Private Sub OpenLetterCreatorWithAddress(addressArray As Variant, rowNumber As Long)
    ' Procedure to open the form with pre-filled address data
    On Error GoTo OpenFormError
    
    ' Check if the form is already open
    Dim existingForm As Object
    On Error Resume Next
    Set existingForm = VBA.UserForms("frmLetterCreator")
    On Error GoTo OpenFormError
    
    If Not existingForm Is Nothing Then
        ' If the form is already open, fill it with new data
        FillFormWithAddressData existingForm, addressArray, rowNumber
        existingForm.SetFocus
        existingForm.ZOrder 0
        Debug.Print "Existing form updated with data from row " & rowNumber
    Else
        ' Load a new form
        Load frmLetterCreator
        FillFormWithAddressData frmLetterCreator, addressArray, rowNumber
        frmLetterCreator.Show vbModeless
        Debug.Print "New form created with data from row " & rowNumber
    End If
    
    Exit Sub
    
OpenFormError:
    MsgBox "Error opening letter creation form: " & Err.description, vbCritical, "Error"
End Sub

Private Sub FillFormWithAddressData(targetForm As Object, addressArray As Variant, rowNumber As Long)
    On Error GoTo FillFormError

    targetForm.LoadAddressForEditing addressArray, rowNumber
    Debug.Print "Form fields filled for editing addressee: " & CStr(addressArray(0))
    Exit Sub

FillFormError:
    MsgBox "Error filling the letter creation form: " & Err.description, vbCritical, "Error"
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' Open form only on double click
    On Error GoTo DoubleClickError
    
    Dim addressRow As Long
    
    If Me.Name = "Addresses" And Target.Cells.count = 1 Then
        addressRow = Target.Row
        
        If addressRow > 1 And Len(Trim(Me.Cells(addressRow, 1).value)) > 0 Then
            Cancel = True  ' Cancel standard double-click behavior
            
            ' Gather data
            Dim addressData(7) As String
            Dim i As Integer
            
            For i = 0 To 7
                addressData(i) = CStr(Me.Cells(addressRow, i + 1).value)
            Next i
            
            ' Open the form
            OpenLetterCreatorWithAddress addressData, addressRow
            
            ' FIXED: Using the new SetStatusBarMessage procedure
            SetStatusBarMessage "Form opened for: " & addressData(0), 3
        End If
    End If
    
    Exit Sub
    
DoubleClickError:
    MsgBox "Double click error: " & Err.description, vbCritical
End Sub