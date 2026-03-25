Attribute VB_Name = "mdlInicialize"
' ======================================================================
' Module: mdlInitialize
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Creation and configuration of Excel worksheets structure for the program
' Version: 1.4.1 — 11.09.2025
' ======================================================================
Option Explicit

Public Sub InitializeAllWorksheets()
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
        .Cells(1, 1).Value = "Recipient Name"
        .Cells(1, 2).Value = "Street"
        .Cells(1, 3).Value = "City"
        .Cells(1, 4).Value = "District"
        .Cells(1, 5).Value = "Region"
        .Cells(1, 6).Value = "Postal Code"
        .Cells(1, 7).Value = "Phone"
        
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
        .Cells(1, 1).Value = "Addressee"
        .Cells(1, 2).Value = "Outgoing Number"
        .Cells(1, 3).Value = "Outgoing Date"
        .Cells(1, 4).Value = "Attachment Name"
        .Cells(1, 5).Value = "Document Sum"
        .Cells(1, 6).Value = "Return Mark"
        .Cells(1, 7).Value = "Executor Name"
        .Cells(1, 8).Value = "Send Type"
        
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
        .Cells(1, 1).Value = "Attachments"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Interior.ColorIndex = 35  ' Light green
        
        ' Examples of attachment types
        .Cells(2, 1).Value = "Outgoing notice"
        .Cells(3, 1).Value = "Material acceptance certificate"
        .Cells(4, 1).Value = "Transfer of FA, IA, NPA"
        .Cells(5, 1).Value = "Invoice"
        .Cells(6, 1).Value = "Waybill"
        .Cells(7, 1).Value = "Certificate of completion"
        
        ' Executors table
        .Cells(1, 3).Value = "Executor Name"
        .Cells(1, 4).Value = "Phone"
        .Cells(1, 3).Font.Bold = True
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 3).Interior.ColorIndex = 36  ' Light yellow
        .Cells(1, 4).Interior.ColorIndex = 36  ' Light yellow
        
        ' Examples of executors
        .Cells(2, 3).Value = "Kerzhaev E.A."
        .Cells(2, 4).Value = "8-928-123-45-67"
        .Cells(3, 3).Value = "Ivanov I.I."
        .Cells(3, 4).Value = "8-928-234-56-78"
        .Cells(4, 3).Value = "Petrov P.P."
        .Cells(4, 4).Value = "8-928-345-67-89"
        
        ' Letter texts table
        .Cells(1, 6).Value = "Text"
        .Cells(1, 6).Font.Bold = True
        .Cells(1, 6).Interior.ColorIndex = 34  ' Light pink
        
        ' Text examples (CHANGE: first letter is lowercase)
        .Cells(2, 6).Value = "forwarding the following documents to your address for confirmation"
        .Cells(3, 6).Value = "forwarding confirmed accounting documents to your address"
        
        ' Creating a structured table in column F
        Dim textRange As Range
        Set textRange = .Range("F1:F3")
        
        On Error Resume Next
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, textRange, , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = "Text"
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
        InitializeAllWorksheets
    End If
End Sub
