Attribute VB_Name = "ModuleAuditLogger"
' ======================================================================
' Module: ModuleAuditLogger
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: User action audit system
' Version: 1.0.0 — 06.08.2025
' ======================================================================

Option Explicit

' Constants for action types
Public Const AUDIT_OPEN_FILE1 As String = "OPEN_FILE"
Public Const AUDIT_CREATE_LETTER As String = "CREATE_LETTER"
Public Const AUDIT_CLOSE_FILE As String = "CLOSE_FILE"
Public Const AUDIT_SEARCH_ADDRESS As String = "SEARCH_ADDRESS"
Public Const AUDIT_SEARCH_ATTACHMENT As String = "SEARCH_ATTACHMENT"
Public Const AUDIT_SAVE_ADDRESS As String = "SAVE_ADDRESS"

Public Sub WriteAuditLog(action As String, details As String, Optional recipient As String = "")
    On Error GoTo AuditError
    
    Dim auditSheet As Worksheet
    Set auditSheet = GetOrCreateAuditSheet()
    
    Dim lastRow As Long
    lastRow = auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row + 1
    
    ' Getting user and computer information
    Dim userName As String, computerName As String, ipAddress As String
    userName = Environ("USERNAME")
    computerName = Environ("COMPUTERNAME")
    ipAddress = GetLocalIPAddress()
    
    With auditSheet
        .Cells(lastRow, 1).Value = Format(Now, "dd.mm.yyyy")           ' Date
        .Cells(lastRow, 2).Value = Format(Now, "hh:mm:ss")             ' Time
        .Cells(lastRow, 3).Value = userName                            ' User
        .Cells(lastRow, 4).Value = computerName                        ' Computer
        .Cells(lastRow, 5).Value = ipAddress                           ' IP Address
        .Cells(lastRow, 6).Value = action                              ' Action
        .Cells(lastRow, 7).Value = details                             ' Details
        .Cells(lastRow, 8).Value = recipient                           ' Letter Recipient
        .Cells(lastRow, 9).Value = Application.Version                 ' Excel Version
        
        ' Color highlighting by action type
        Select Case action
            Case AUDIT_OPEN_FILE1
                .Cells(lastRow, 6).Interior.Color = RGB(200, 255, 200)  ' Light green
            Case AUDIT_CREATE_LETTER
                .Cells(lastRow, 6).Interior.Color = RGB(255, 255, 200)  ' Yellow
            Case AUDIT_CLOSE_FILE
                .Cells(lastRow, 6).Interior.Color = RGB(255, 200, 200)  ' Light red
            Case Else
                .Cells(lastRow, 6).Interior.Color = RGB(240, 240, 240)  ' Gray
        End Select
    End With
    
    ' Cleaning up old records (older than 90 days for audit)
    CleanOldAuditEntries auditSheet, 90
    
    Exit Sub
    
AuditError:
    ' Critical error - write to Debug
    Debug.Print Format(Now, "hh:mm:ss") & " AUDIT_ERROR: " & Err.description
End Sub

Private Function GetOrCreateAuditSheet() As Worksheet
    Dim auditSheet As Worksheet
    
    ' Attempting to find the "AuditLog" sheet
    On Error Resume Next
    Set auditSheet = ThisWorkbook.Worksheets("AuditLog")
    On Error GoTo 0
    
    ' Create sheet if not found
    If auditSheet Is Nothing Then
        Set auditSheet = ThisWorkbook.Worksheets.Add
        With auditSheet
            .Name = "AuditLog"
            .Visible = xlSheetVeryHidden  ' Hide the sheet
            
            ' Headers
            .Cells(1, 1).Value = "Date"
            .Cells(1, 2).Value = "Time"
            .Cells(1, 3).Value = "User"
            .Cells(1, 4).Value = "Computer"
            .Cells(1, 5).Value = "IP Address"
            .Cells(1, 6).Value = "Action"
            .Cells(1, 7).Value = "Details"
            .Cells(1, 8).Value = "Recipient"
            .Cells(1, 9).Value = "Excel Version"
            
            ' Headers formatting
            With .Range("A1:I1")
                .Font.Bold = True
                .Interior.Color = RGB(100, 100, 100)
                .Font.Color = RGB(255, 255, 255)
                .Borders.LineStyle = xlContinuous
            End With
            
            .Columns("A:I").AutoFit
        End With
    End If
    
    Set GetOrCreateAuditSheet = auditSheet
End Function

Private Function GetLocalIPAddress() As String
    On Error Resume Next
    ' Simple IP retrieval via WMI
    Dim objWMIService As Object
    Dim colItems As Object, objItem As Object
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    
    For Each objItem In colItems
        If Not IsNull(objItem.ipAddress) Then
            GetLocalIPAddress = objItem.ipAddress(0)
            Exit For
        End If
    Next
    
    If GetLocalIPAddress = "" Then GetLocalIPAddress = "Unknown"
    On Error GoTo 0
End Function

Private Sub CleanOldAuditEntries(auditSheet As Worksheet, daysToKeep As Integer)
    On Error Resume Next
    
    Dim lastRow As Long
    lastRow = auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row
    
    If lastRow <= 5000 Then Exit Sub  ' Clean only if there are more than 5000 records
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim logDate As Date
        If IsDate(auditSheet.Cells(i, 1).Value) Then
            logDate = CDate(auditSheet.Cells(i, 1).Value)
            
            ' Delete records older than the specified number of days
            If Date - logDate > daysToKeep Then
                auditSheet.Rows(i).Delete
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

Public Sub ShowAuditLog()
    ' Function for administrator - show audit log
    Dim auditSheet As Worksheet
    Set auditSheet = GetOrCreateAuditSheet()
    
    auditSheet.Visible = xlSheetVisible
    auditSheet.Activate
    MsgBox "Audit log opened. Do not forget to hide the sheet after viewing!", vbInformation
End Sub

Public Sub GenerateAuditReport(Optional daysBack As Integer = 30)
    ' Generate usage report for a period
    On Error GoTo ReportError
    
    Dim auditSheet As Worksheet
    Set auditSheet = GetOrCreateAuditSheet()
    
    Dim reportWb As Workbook
    Dim reportWs As Worksheet
    Set reportWb = Workbooks.Add
    Set reportWs = reportWb.Worksheets(1)
    
    reportWs.Name = "Audit report for " & daysBack & " days"
    
    ' Report headers
    With reportWs
        .Cells(1, 1).Value = "AUDIT REPORT FOR 'LETTER GENERATION' SYSTEM"
        .Cells(2, 1).Value = "Period: " & Format(Date - daysBack, "dd.mm.yyyy") & " - " & Format(Date, "dd.mm.yyyy")
        .Cells(3, 1).Value = "Report generation date: " & Format(Now, "dd.mm.yyyy hh:mm")
        
        .Cells(5, 1).Value = "Date"
        .Cells(5, 2).Value = "Time"
        .Cells(5, 3).Value = "User"
        .Cells(5, 4).Value = "Computer"
        .Cells(5, 5).Value = "Action"
        .Cells(5, 6).Value = "Details"
        .Cells(5, 7).Value = "Recipient"
        
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A5:G5").Font.Bold = True
        .Range("A5:G5").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Copying data for the period
    Dim sourceRow As Long, targetRow As Long
    targetRow = 6
    
    For sourceRow = 2 To auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row
        If IsDate(auditSheet.Cells(sourceRow, 1).Value) Then
            Dim entryDate As Date
            entryDate = CDate(auditSheet.Cells(sourceRow, 1).Value)
            
            If Date - entryDate <= daysBack Then
                reportWs.Cells(targetRow, 1).Value = auditSheet.Cells(sourceRow, 1).Value  ' Date
                reportWs.Cells(targetRow, 2).Value = auditSheet.Cells(sourceRow, 2).Value  ' Time
                reportWs.Cells(targetRow, 3).Value = auditSheet.Cells(sourceRow, 3).Value  ' User
                reportWs.Cells(targetRow, 4).Value = auditSheet.Cells(sourceRow, 4).Value  ' Computer
                reportWs.Cells(targetRow, 5).Value = auditSheet.Cells(sourceRow, 6).Value  ' Action
                reportWs.Cells(targetRow, 6).Value = auditSheet.Cells(sourceRow, 7).Value  ' Details
                reportWs.Cells(targetRow, 7).Value = auditSheet.Cells(sourceRow, 8).Value  ' Recipient
                targetRow = targetRow + 1
            End If
        End If
    Next sourceRow
    
    reportWs.Columns("A:G").AutoFit
    reportWb.Application.Visible = True
    
    Exit Sub
    
ReportError:
    MsgBox "Error generating report: " & Err.description, vbCritical
End Sub




' Show usage statistics
Public Sub ShowUsageStatistics()
    Dim auditSheet As Worksheet
    Set auditSheet = GetOrCreateAuditSheet()
    
    Dim totalSessions As Long, totalLetters As Long
    Dim uniqueUsers As Object
    Set uniqueUsers = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row
        Dim action As String, user As String
        action = auditSheet.Cells(i, 6).Value
        user = auditSheet.Cells(i, 3).Value
        
        If action = AUDIT_OPEN_FILE1 Then totalSessions = totalSessions + 1
        If action = AUDIT_CREATE_LETTER Then totalLetters = totalLetters + 1
        
        If Not uniqueUsers.Exists(user) Then uniqueUsers.Add user, 0
        uniqueUsers(user) = uniqueUsers(user) + 1
    Next i
    
    Dim report As String
    report = "SYSTEM USAGE STATISTICS" & vbCrLf & vbCrLf
    report = report & "Total sessions: " & totalSessions & vbCrLf
    report = report & "Letters created: " & totalLetters & vbCrLf
    report = report & "Unique users: " & uniqueUsers.count & vbCrLf & vbCrLf
    report = report & "TOP USERS:" & vbCrLf
    
    Dim key As Variant
    For Each key In uniqueUsers.keys
        report = report & key & ": " & uniqueUsers(key) & " actions" & vbCrLf
    Next key
    
    MsgBox report, vbInformation, "System Statistics"
End Sub

' Quick commands for the administrator
Public Sub AdminPanel()
    Dim choice As String
    choice = InputBox("Select action:" & vbCrLf & _
                     "1 - Show audit log" & vbCrLf & _
                     "2 - Usage statistics" & vbCrLf & _
                     "3 - 30-day report" & vbCrLf & _
                     "4 - 7-day report", "Administrator Panel", "1")
    
    Select Case choice
        Case "1": ShowAuditLog
        Case "2": ShowUsageStatistics
        Case "3": GenerateAuditReport 30
        Case "4": GenerateAuditReport 7
        Case Else: MsgBox "Cancelled", vbInformation
    End Select
End Sub

