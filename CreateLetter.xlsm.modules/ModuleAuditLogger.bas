Attribute VB_Name = "ModuleAuditLogger"
' ======================================================================
' Module: ModuleAuditLogger
' Author: CreateLetter contributors
' Purpose: Audit log helpers for workbook activity tracking
' Version: 1.0.5 - 29.03.2026
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
    lastRow = auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row + 1

    userName = Environ$("USERNAME")
    computerName = Environ$("COMPUTERNAME")
    ipAddress = GetLocalIPAddress()

    With auditSheet
        .Cells(lastRow, 1).value = Format$(Now, "dd.mm.yyyy")
        .Cells(lastRow, 2).value = Format$(Now, "hh:mm:ss")
        .Cells(lastRow, 3).value = userName
        .Cells(lastRow, 4).value = computerName
        .Cells(lastRow, 5).value = ipAddress
        .Cells(lastRow, 6).value = action
        .Cells(lastRow, 7).value = details
        .Cells(lastRow, 8).value = recipient
        .Cells(lastRow, 9).value = Application.Version

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
    Debug.Print Format$(Now, "hh:mm:ss") & " AUDIT_ERROR: " & Err.description
End Sub

Private Function GetOrCreateAuditSheet() As Worksheet
    Dim auditSheet As Worksheet

    Set auditSheet = TryGetWorksheetByName("AuditLog")

    If auditSheet Is Nothing Then
        Set auditSheet = ThisWorkbook.Worksheets.Add
        With auditSheet
            .Name = "AuditLog"
            .Visible = xlSheetVeryHidden
            .Cells(1, 1).value = "Date"
            .Cells(1, 2).value = "Time"
            .Cells(1, 3).value = "User"
            .Cells(1, 4).value = "Computer"
            .Cells(1, 5).value = "IP Address"
            .Cells(1, 6).value = "Action"
            .Cells(1, 7).value = "Details"
            .Cells(1, 8).value = "Recipient"
            .Cells(1, 9).value = "Excel Version"

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
    On Error GoTo LookupError

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")

    For Each objItem In colItems
        If Not IsNull(objItem.ipAddress) Then
            GetLocalIPAddress = objItem.ipAddress(0)
            Exit For
        End If
    Next objItem

    If GetLocalIPAddress = "" Then GetLocalIPAddress = "Unknown"
    Exit Function

LookupError:
    GetLocalIPAddress = "Unknown"
    Debug.Print "GetLocalIPAddress error: " & Err.description
End Function

Private Sub CleanOldAuditEntries(auditSheet As Worksheet, daysToKeep As Integer)
    On Error GoTo CleanupError

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim logDate As Date

    lastRow = auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row
    If lastRow <= 5000 Then Exit Sub

    For rowIndex = lastRow To 2 Step -1
        If IsDate(auditSheet.Cells(rowIndex, 1).value) Then
            logDate = CDate(auditSheet.Cells(rowIndex, 1).value)
            If Date - logDate > daysToKeep Then
                auditSheet.Rows(rowIndex).Delete
            End If
        End If
    Next rowIndex
    Exit Sub

CleanupError:
    Debug.Print "CleanOldAuditEntries error: " & Err.description
End Sub

Private Function TryGetWorksheetByName(sheetName As String) As Worksheet
    On Error GoTo LookupError

    Set TryGetWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    Exit Function

LookupError:
    Set TryGetWorksheetByName = Nothing
End Function

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
        .Cells(1, 1).value = "AUDIT REPORT FOR LETTER SYSTEM"
        .Cells(2, 1).value = "Period: " & Format$(Date - daysBack, "dd.mm.yyyy") & " - " & Format$(Date, "dd.mm.yyyy")
        .Cells(3, 1).value = "Generated: " & Format$(Now, "dd.mm.yyyy hh:mm")
        .Cells(5, 1).value = "Date"
        .Cells(5, 2).value = "Time"
        .Cells(5, 3).value = "User"
        .Cells(5, 4).value = "Computer"
        .Cells(5, 5).value = "Action"
        .Cells(5, 6).value = "Details"
        .Cells(5, 7).value = "Recipient"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A5:G5").Font.Bold = True
        .Range("A5:G5").Interior.Color = RGB(200, 200, 200)
    End With

    targetRow = 6
    For sourceRow = 2 To auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row
        If IsDate(auditSheet.Cells(sourceRow, 1).value) Then
            entryDate = CDate(auditSheet.Cells(sourceRow, 1).value)
            If Date - entryDate <= daysBack Then
                reportWs.Cells(targetRow, 1).value = auditSheet.Cells(sourceRow, 1).value
                reportWs.Cells(targetRow, 2).value = auditSheet.Cells(sourceRow, 2).value
                reportWs.Cells(targetRow, 3).value = auditSheet.Cells(sourceRow, 3).value
                reportWs.Cells(targetRow, 4).value = auditSheet.Cells(sourceRow, 4).value
                reportWs.Cells(targetRow, 5).value = auditSheet.Cells(sourceRow, 6).value
                reportWs.Cells(targetRow, 6).value = auditSheet.Cells(sourceRow, 7).value
                reportWs.Cells(targetRow, 7).value = auditSheet.Cells(sourceRow, 8).value
                targetRow = targetRow + 1
            End If
        End If
    Next sourceRow

    reportWs.Columns("A:G").AutoFit
    reportWb.Application.Visible = True
    Exit Sub

ReportError:
    MsgBox t("audit.msg.report_error", "Error generating audit report: ") & Err.description, vbCritical
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

    For rowIndex = 2 To auditSheet.Cells(auditSheet.Rows.count, 1).End(xlUp).Row
        action = auditSheet.Cells(rowIndex, 6).value
        userName = auditSheet.Cells(rowIndex, 3).value

        If action = AUDIT_OPEN_FILE1 Then totalSessions = totalSessions + 1
        If action = AUDIT_CREATE_LETTER Then totalLetters = totalLetters + 1

        If Not uniqueUsers.Exists(userName) Then uniqueUsers.Add userName, 0
        uniqueUsers(userName) = uniqueUsers(userName) + 1
    Next rowIndex

    report = t("audit.msg.system_statistics_header", "SYSTEM USAGE STATISTICS") & vbCrLf & vbCrLf
    report = report & t("audit.msg.total_sessions", "Total sessions: ") & totalSessions & vbCrLf
    report = report & t("audit.msg.letters_created", "Letters created: ") & totalLetters & vbCrLf
    report = report & t("audit.msg.unique_users", "Unique users: ") & uniqueUsers.count & vbCrLf & vbCrLf
    report = report & t("audit.msg.top_users", "TOP USERS:") & vbCrLf

    For Each key In uniqueUsers.keys
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

