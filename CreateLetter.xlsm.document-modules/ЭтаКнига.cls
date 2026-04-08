VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ======================================================================
' Workbook Module: ThisWorkbook
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Handling Excel workbook events + VbaModuleManager
' Version: 1.5.0 — 16.12.2025
' ======================================================================

Option Explicit

' ----------------------------------------------------------------------
' WORKBOOK EVENTS
' ----------------------------------------------------------------------

Private Sub Workbook_Open()
    ' 1. First, import the updated code (VbaModuleManager)
    ' If you are developing, leave this line.
    ' If you are giving the ready version to a user without the src folder, comment it out.
'   On Error Resume Next ' So the import error doesn't block your program from starting
'   Call ImportModules(ThisWorkbook.Name & ".modules", ShowMsgBox:=True)
'   On Error GoTo 0
    
    ' 2. Main program logic
    ' Check for the presence of necessary sheets when opening the file
    CheckWorksheetStructure
    
    ' Automatically open the letter creation form with a delay
    Application.OnTime Now + TimeValue("00:00:02"), "ShowLetterCreatorDelayed"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    ' Export the code to a folder before saving.
    ' This allows saving module versions in Git or simply having a backup of the code.
    ' Comment out before giving to users if they don't need it.
'   On Error Resume Next
'   Call ExportModules(ThisWorkbook.Name & ".modules")
'   On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    ' Clear modules before closing.
    ' WARNING: If this is enabled, the next time you open the workbook it will be empty (no code),
    ' and it will be forced to load the code from the folder via Workbook_Open.
    ' This is usually needed to keep a "clean" Excel file in a repository.
    ' If you are afraid of losing the code, it's better to leave this line commented out until you understand it.
    
    ' Call RemoveModules(ShowMsgBox:=True)
End Sub

' ----------------------------------------------------------------------
' AUXILIARY PROCEDURES
' ----------------------------------------------------------------------

Private Sub CheckWorksheetStructure()
    Dim hasAddresses As Boolean, hasLetters As Boolean, hasSettings As Boolean
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "Addresses"
                hasAddresses = True
            Case "Letters"
                hasLetters = True
            Case "Settings"
                hasSettings = True
        End Select
    Next ws
    
    If Not (hasAddresses And hasLetters And hasSettings) Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Required sheets not found. Create data structure?", vbYesNo + vbQuestion)
        
        If response = vbYes Then
            ' Use the function from the new module (ensure the module with this function is loaded)
            ' InitializeAllWorksheets ' <-- Uncomment if this function is available
        End If
    End If
End Sub

Private Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Main program launch procedure (for manual call)
Public Sub StartFormirovanieLetters()
    On Error GoTo StartupError
    
    ' Check if the form is already open
    Dim existingForm As Object
    On Error Resume Next
    ' Access the UserForms collection to check if the form is loaded
    Dim i As Integer
    For i = 0 To VBA.UserForms.count - 1
        If VBA.UserForms(i).Name = "frmLetterCreator" Then
            Set existingForm = VBA.UserForms(i)
            Exit For
        End If
    Next i
    On Error GoTo StartupError
    
    If Not existingForm Is Nothing Then
        existingForm.SetFocus
        ' ZOrder doesn't always work with UserForm, but if it worked for you - ok
        ' existingForm.ZOrder 0
        Debug.Print "Letter creation form is already open"
    Else
        ' Launch the form
        ' Ensure that frmLetterCreator exists in the project
        frmLetterCreator.Show vbModeless
        Debug.Print "Letter creation form launched modelessly"
    End If
    
    Exit Sub
    
StartupError:
    MsgBox "Error opening letter creation form: " & Err.description, vbCritical
End Sub