Attribute VB_Name = "mdlInicialize"
' ======================================================================
' Module: mdlInitialize
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Workbook sheet bootstrap and reset entry points with English-safe public aliases
' Version: 1.4.5 - 27.03.2026
' ======================================================================
Option Explicit

Public Sub InitializeAllWorksheets()
    BootstrapWorkbookSheets
End Sub

Public Sub BootstrapWorkbookSheets()
    ' Creation and configuration of all necessary sheets (renamed to avoid conflict)
    On Error GoTo InitError
    
    ' "Addresses" sheet
    CreateAddressesSheet
    
    ' "Letters" sheet
    CreateLettersSheet
    
    ' "Settings" sheet
    CreateSettingsSheet
    
    MsgBox "Sheets structure created successfully!", vbInformation
    Exit Sub
    
InitError:
    MsgBox "Error creating sheets structure: " & Err.description, vbCritical
End Sub

Private Sub CreateAddressesSheet()
    Dim ws As Worksheet
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Addresses")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Addresses"
    End If
    
    ' Create headers (CHANGE: added 7th column "Phone")
    With ws
        .Cells(1, 1).value = "Recipient Name"
        .Cells(1, 2).value = "Street"
        .Cells(1, 3).value = "City"
        .Cells(1, 4).value = "District"
        .Cells(1, 5).value = "Region"
        .Cells(1, 6).value = "Postal Code"
        .Cells(1, 7).value = "Phone"
        
        ' Headers formatting (using standard VBA constants)
        With .Range("A1:G1")
            .Font.Bold = True
            .Interior.ColorIndex = 37  ' Light blue
            .EntireColumn.AutoFit
        End With
    End With
End Sub

Private Sub CreateLettersSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Letters")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Letters"
    End If
    
    ' Create headers
    With ws
        .Cells(1, 1).value = "Addressee"
        .Cells(1, 2).value = "Outgoing Number"
        .Cells(1, 3).value = "Outgoing Date"
        .Cells(1, 4).value = "Attachment Name"
        .Cells(1, 5).value = "Document Sum"
        .Cells(1, 6).value = "Return Mark"
        .Cells(1, 7).value = "Executor Name"
        .Cells(1, 8).value = "Send Type"
        
        ' Headers formatting
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.ColorIndex = 40  ' Light orange
            .EntireColumn.AutoFit
        End With
    End With
End Sub

Private Sub CreateSettingsSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Settings"
    End If
    
    ' Attachment types table
    With ws
        .Cells(1, 1).value = "Attachments"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Interior.ColorIndex = 35  ' Light green
        
        ' Examples of attachment types
        .Cells(2, 1).value = "Outgoing notice"
        .Cells(3, 1).value = "Material acceptance certificate"
        .Cells(4, 1).value = "Transfer of FA, IA, NPA"
        .Cells(5, 1).value = "Invoice"
        .Cells(6, 1).value = "Waybill"
        .Cells(7, 1).value = "Certificate of completion"
        
        ' Executors table
        .Cells(1, 3).value = "Executor Name"
        .Cells(1, 4).value = "Phone"
        .Cells(1, 3).Font.Bold = True
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 3).Interior.ColorIndex = 36  ' Light yellow
        .Cells(1, 4).Interior.ColorIndex = 36  ' Light yellow
        
        ' Examples of executors
        .Cells(2, 3).value = "Kerzhaev E.A."
        .Cells(2, 4).value = "8-928-123-45-67"
        .Cells(3, 3).value = "Ivanov I.I."
        .Cells(3, 4).value = "8-928-234-56-78"
        .Cells(4, 3).value = "Petrov P.P."
        .Cells(4, 4).value = "8-928-345-67-89"
        
        ' Letter texts table
        .Cells(1, 6).value = "Text"
        .Cells(1, 6).Font.Bold = True
        .Cells(1, 6).Interior.ColorIndex = 34  ' Light pink
        
        ' Text examples (CHANGE: first letter is lowercase)
        .Cells(2, 6).value = "forwarding the following documents to your address for confirmation"
        .Cells(3, 6).value = "forwarding confirmed accounting documents to your address"
        
        ' Creating a structured table in column F
        Dim textRange As Range
        Set textRange = .Range("F1:F3")
        
        On Error Resume Next
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, textRange, , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = "tblLetterTexts"
        End If
        On Error GoTo 0
        
        ' Auto-fit column widths
        .Columns("A:F").AutoFit
    End With
End Sub

Public Sub ResetWorksheets()
    ' Reset all worksheet settings
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to reset all data?", vbYesNo + vbQuestion + vbDefaultButton2)
    
    If response = vbYes Then
        On Error Resume Next
        
        ' Delete existing sheets
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Addresses").Delete
        ThisWorkbook.Worksheets("Letters").Delete
        ThisWorkbook.Worksheets("Settings").Delete
        Application.DisplayAlerts = True
        
        On Error GoTo 0
        
        ' Recreate
        BootstrapWorkbookSheets
    End If
End Sub

Public Sub ResetWorkbookSheets()
    ResetWorksheets
End Sub

