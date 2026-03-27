Attribute VB_Name = "ModuleAuditLogger"
' ======================================================================
' Module: ModuleAuditLogger
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Audit log helpers for workbook activity tracking
' Version: 1.0.3 - 27.03.2026
' ======================================================================

Option Explicit

Public Const AUDIT_OPEN_FILE1 As String = "OPEN_FILE"
Public Const AUDIT_CREATE_LETTER As String = "CREATE_LETTER"
Public Const AUDIT_CLOSE_FILE As String = "CLOSE_FILE"
Public Const AUDIT_SEARCH_ADDRESS As String = "SEARCH_ADDRESS"
Public Const AUDIT_SEARCH_ATTACHMENT As String = "SEARCH_ATTACHMENT"
Public Const AUDIT_SAVE_ADDRESS As String = "SAVE_ADDRESS"

Public Sub OpenAuditLog()
    ShowAuditLog
End Sub

Public Sub WriteAuditLog(action As String, details As String, Optional recipient As String = "")
    On Error GoTo AuditError

    Dim auditSheet As Worksheet
    Dim lastRow As Long
    Dim userName As String
    Dim computerName As String
    Dim ipAddress As String

    Set auditSheet = GetOrCreateAuditSheet()
    lastRow = auditSheet.Cells(auditSheet.Rows.Count, 1).End(xlUp).Row + 1

    userName = Environ$("USERNAME")
    computerName = Environ$("COMPUTERNAME")
    ipAddress = GetLocalIPAddress()

    With auditSheet
        .Cells(lastRow, 1).Value = Format$(Now, "dd.mm.yyyy")
        .Cells(lastRow, 2).Value = Format$(Now, "hh:mm:ss")
        .Cells(lastRow, 3).Value = userName
        .Cells(lastRow, 4).Value = computerName
        .Cells(lastRow, 5).Value = ipAddress
        .Cells(lastRow, 6).Value = action
        .Cells(lastRow, 7).Value = details
        .Cells(lastRow, 8).Value = recipient
        .Cells(lastRow, 9).Value = Application.Version

        Select Case action
            Case AUDIT_OPEN_FILE1
                .Cells(lastRow, 6).Interior.Color = RGB(200, 255, 200)
            Case AUDIT_CREATE_LETTER
                .Cells(lastRow, 6).Interior.Color = RGB(255, 255, 200)
            Case AUDIT_CLOSE_FILE
                .Cells(lastRow, 6).Interior.Color = RGB(255, 200, 200)
            Case Else
                .Cells(lastRow, 6).Interior.Color = RGB(240, 240, 240)
        End Select
    End With

    CleanOldAuditEntries auditSheet, 90
    Exit Sub

AuditError:
    Debug.Print Format$(Now, "hh:mm:ss") & " AUDIT_ERROR: " & Err.Description
End Sub

Private Function GetOrCreateAuditSheet() As Worksheet
    Dim auditSheet As Worksheet

    On Error Resume Next
    Set auditSheet = ThisWorkbook.Worksheets("AuditLog")
    On Error GoTo 0

    If auditSheet Is Nothing Then
        Set auditSheet = ThisWorkbook.Worksheets.Add
        With auditSheet
            .Name = "AuditLog"
            .Visible = xlSheetVeryHidden
            .Cells(1, 1).Value = "Date"
            .Cells(1, 2).Value = "Time"
            .Cells(1, 3).Value = "User"
            .Cells(1, 4).Value = "Computer"
            .Cells(1, 5).Value = "IP Address"
            .Cells(1, 6).Value = "Action"
            .Cells(1, 7).Value = "Details"
            .Cells(1, 8).Value = "Recipient"
            .Cells(1, 9).Value = "Excel Version"

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

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")

    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then
            GetLocalIPAddress = objItem.IPAddress(0)
            Exit For
        End If
    Next objItem

    If GetLocalIPAddress = "" Then GetLocalIPAddress = "Unknown"
    On Error GoTo 0
End Function

Private Sub CleanOldAuditEntries(auditSheet As Worksheet, daysToKeep As Integer)
    On Error Resume Next

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim logDate As Date

    lastRow = auditSheet.Cells(auditSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 5000 Then Exit Sub

    For rowIndex = lastRow To 2 Step -1
        If IsDate(auditSheet.Cells(rowIndex, 1).Value) Then
            logDate = CDate(auditSheet.Cells(rowIndex, 1).Value)
            If Date - logDate > daysToKeep Then
                auditSheet.Rows(rowIndex).Delete
            End If
        End If
    Next rowIndex

    On Error GoTo 0
End Sub

Public Sub ShowAuditLog()
    Dim auditSheet As Worksheet
    Set auditSheet = GetOrCreateAuditSheet()

    auditSheet.Visible = xlSheetVisible
    auditSheet.Activate
    MsgBox t("audit.msg.log_opened", "Audit log opened. Hide the sheet after review if needed."), vbInformation
End Sub

Public Sub GenerateAuditReport(Optional daysBack As Integer = 30)
    On Error GoTo ReportError

    Dim auditSheet As Worksheet
    Dim reportWb As Workbook
    Dim reportWs As Worksheet
    Dim sourceRow As Long
    Dim targetRow As Long
    Dim entryDate As Date

    Set auditSheet = GetOrCreateAuditSheet()
    Set reportWb = Workbooks.Add
    Set reportWs = reportWb.Worksheets(1)

    reportWs.Name = "Audit report " & daysBack & " days"

    With reportWs
        .Cells(1, 1).Value = "AUDIT REPORT FOR LETTER SYSTEM"
        .Cells(2, 1).Value = "Period: " & Format$(Date - daysBack, "dd.mm.yyyy") & " - " & Format$(Date, "dd.mm.yyyy")
        .Cells(3, 1).Value = "Generated: " & Format$(Now, "dd.mm.yyyy hh:mm")
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

    targetRow = 6
    For sourceRow = 2 To auditSheet.Cells(auditSheet.Rows.Count, 1).End(xlUp).Row
        If IsDate(auditSheet.Cells(sourceRow, 1).Value) Then
            entryDate = CDate(auditSheet.Cells(sourceRow, 1).Value)
            If Date - entryDate <= daysBack Then
                reportWs.Cells(targetRow, 1).Value = auditSheet.Cells(sourceRow, 1).Value
                reportWs.Cells(targetRow, 2).Value = auditSheet.Cells(sourceRow, 2).Value
                reportWs.Cells(targetRow, 3).Value = auditSheet.Cells(sourceRow, 3).Value
                reportWs.Cells(targetRow, 4).Value = auditSheet.Cells(sourceRow, 4).Value
                reportWs.Cells(targetRow, 5).Value = auditSheet.Cells(sourceRow, 6).Value
                reportWs.Cells(targetRow, 6).Value = auditSheet.Cells(sourceRow, 7).Value
                reportWs.Cells(targetRow, 7).Value = auditSheet.Cells(sourceRow, 8).Value
                targetRow = targetRow + 1
            End If
        End If
    Next sourceRow

    reportWs.Columns("A:G").AutoFit
    reportWb.Application.Visible = True
    Exit Sub

ReportError:
    MsgBox t("audit.msg.report_error", "Error generating audit report: ") & Err.Description, vbCritical
End Sub

Public Sub ShowAuditStatistics()
    ShowUsageStatistics
End Sub

Public Sub ShowUsageStatistics()
    Dim auditSheet As Worksheet
    Dim totalSessions As Long
    Dim totalLetters As Long
    Dim uniqueUsers As Object
    Dim rowIndex As Long
    Dim action As String
    Dim userName As String
    Dim report As String
    Dim key As Variant

    Set auditSheet = GetOrCreateAuditSheet()
    Set uniqueUsers = CreateObject("Scripting.Dictionary")

    For rowIndex = 2 To auditSheet.Cells(auditSheet.Rows.Count, 1).End(xlUp).Row
        action = auditSheet.Cells(rowIndex, 6).Value
        userName = auditSheet.Cells(rowIndex, 3).Value

        If action = AUDIT_OPEN_FILE1 Then totalSessions = totalSessions + 1
        If action = AUDIT_CREATE_LETTER Then totalLetters = totalLetters + 1

        If Not uniqueUsers.Exists(userName) Then uniqueUsers.Add userName, 0
        uniqueUsers(userName) = uniqueUsers(userName) + 1
    Next rowIndex

    report = t("audit.msg.system_statistics_header", "SYSTEM USAGE STATISTICS") & vbCrLf & vbCrLf
    report = report & t("audit.msg.total_sessions", "Total sessions: ") & totalSessions & vbCrLf
    report = report & t("audit.msg.letters_created", "Letters created: ") & totalLetters & vbCrLf
    report = report & t("audit.msg.unique_users", "Unique users: ") & uniqueUsers.Count & vbCrLf & vbCrLf
    report = report & t("audit.msg.top_users", "TOP USERS:") & vbCrLf

    For Each key In uniqueUsers.Keys
        report = report & key & ": " & uniqueUsers(key) & " actions" & vbCrLf
    Next key

    MsgBox report, vbInformation, t("audit.title.system_statistics", "System Statistics")
End Sub

Public Sub OpenAuditAdminPanel()
    AdminPanel
End Sub

Public Sub AdminPanel()
    Dim choice As String

    choice = InputBox(t("audit.msg.admin_prompt", "Select action:") & vbCrLf & _
                      t("audit.msg.admin_option_1", "1 - Show audit log") & vbCrLf & _
                      t("audit.msg.admin_option_2", "2 - Usage statistics") & vbCrLf & _
                      t("audit.msg.admin_option_3", "3 - 30-day report") & vbCrLf & _
                      t("audit.msg.admin_option_4", "4 - 7-day report"), _
                      t("audit.title.admin_panel", "Administrator Panel"), "1")

    Select Case choice
        Case "1": ShowAuditLog
        Case "2": ShowUsageStatistics
        Case "3": GenerateAuditReport 30
        Case "4": GenerateAuditReport 7
        Case Else: MsgBox t("audit.msg.cancelled", "Cancelled"), vbInformation
    End Select
End Sub
