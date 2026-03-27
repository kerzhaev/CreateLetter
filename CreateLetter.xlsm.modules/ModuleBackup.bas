Attribute VB_Name = "ModuleBackup"
' ======================================================================
' Module: ModuleBackup
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Workbook backup helpers
' Version: 1.4.2 - 27.03.2026
' ======================================================================

Option Explicit

Public Sub CreateBackup()
    On Error GoTo BackupError

    Dim backupFolder As String
    Dim backupFileName As String
    Dim backupPath As String

    backupFolder = ThisWorkbook.Path & "\Backups\"
    If Dir$(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If

    backupFileName = "FormirovanieLetters_backup_" & Format$(Now, "yyyy-mm-dd_hh-mm-ss") & ".xlsx"
    backupPath = backupFolder & backupFileName

    ThisWorkbook.SaveCopyAs backupPath
    CleanOldBackups backupFolder, 7

    Debug.Print "Backup created: " & backupFileName
    MsgBox "Backup created successfully!" & vbCrLf & backupPath, vbInformation
    Exit Sub

BackupError:
    Debug.Print "Error creating backup: " & Err.Description
    MsgBox "Error creating backup: " & Err.Description, vbCritical
End Sub

Private Sub CleanOldBackups(backupFolder As String, daysToKeep As Integer)
    On Error Resume Next

    Dim fileName As String
    Dim filePath As String
    Dim fileDate As Date

    fileName = Dir$(backupFolder & "FormirovanieLetters_backup_*.xlsx")
    Do While fileName <> ""
        filePath = backupFolder & fileName
        fileDate = FileDateTime(filePath)

        If Date - fileDate > daysToKeep Then
            Kill filePath
            Debug.Print "Old backup deleted: " & fileName
        End If

        fileName = Dir$
    Loop

    On Error GoTo 0
End Sub

Public Sub AutoBackupOnStartup()
    On Error Resume Next

    Dim lastBackupDate As Date
    lastBackupDate = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))

    If Date - lastBackupDate >= 1 Then
        CreateBackup
        SaveSetting "FormirovanieLetters", "Backup", "LastBackupDate", Date
    End If

    On Error GoTo 0
End Sub

Public Sub ShowBackupInfo()
    Dim backupFolder As String
    Dim fileName As String
    Dim backupList As String
    Dim backupCount As Integer
    Dim filePath As String
    Dim fileDate As Date
    Dim fileSize As Long

    backupFolder = ThisWorkbook.Path & "\Backups\"
    If Dir$(backupFolder, vbDirectory) = "" Then
        MsgBox "Backup folder not found.", vbInformation
        Exit Sub
    End If

    fileName = Dir$(backupFolder & "FormirovanieLetters_backup_*.xlsx")
    backupList = "BACKUP LIST:" & vbCrLf & vbCrLf

    Do While fileName <> ""
        filePath = backupFolder & fileName
        fileDate = FileDateTime(filePath)
        fileSize = FileLen(filePath)

        backupList = backupList & fileName & vbCrLf
        backupList = backupList & "  Date: " & Format$(fileDate, "dd.mm.yyyy hh:mm") & vbCrLf
        backupList = backupList & "  Size: " & Format$(fileSize \ 1024, "#,##0") & " KB" & vbCrLf & vbCrLf

        backupCount = backupCount + 1
        fileName = Dir$
    Loop

    If backupCount = 0 Then
        MsgBox "No backups found.", vbInformation
    Else
        backupList = "Backups found: " & backupCount & vbCrLf & vbCrLf & backupList
        MsgBox backupList, vbInformation, "Backup Information"
    End If
End Sub

Public Sub RestoreFromBackup()
    Dim backupFolder As String

    backupFolder = ThisWorkbook.Path & "\Backups\"
    If Dir$(backupFolder, vbDirectory) = "" Then
        MsgBox "Backup folder not found.", vbExclamation
        Exit Sub
    End If

    MsgBox "To restore from a backup:" & vbCrLf & vbCrLf & _
           "1. Close the current workbook" & vbCrLf & _
           "2. Open the folder: " & backupFolder & vbCrLf & _
           "3. Select the desired backup copy" & vbCrLf & _
           "4. Rename and open the restored workbook", _
           vbInformation, "Restore from Backup"
End Sub

Public Function GetBackupFolderPath() As String
    GetBackupFolderPath = ThisWorkbook.Path & "\Backups\"
End Function

Public Function GetLastBackupDate() As Date
    On Error Resume Next
    GetLastBackupDate = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))
    On Error GoTo 0
End Function

Public Sub SetBackupSettings(enableAutoBackup As Boolean, retentionDays As Integer)
    SaveSetting "FormirovanieLetters", "Backup", "AutoBackupEnabled", enableAutoBackup
    SaveSetting "FormirovanieLetters", "Backup", "RetentionDays", retentionDays

    MsgBox "Backup settings saved:" & vbCrLf & _
           "Automatic backup: " & IIf(enableAutoBackup, "Enabled", "Disabled") & vbCrLf & _
           "Keep copies for: " & retentionDays & " days", vbInformation
End Sub

Public Function GetBackupSettings() As String
    On Error Resume Next

    Dim autoEnabled As Boolean
    Dim retentionDays As Integer
    Dim lastBackup As Date

    autoEnabled = GetSetting("FormirovanieLetters", "Backup", "AutoBackupEnabled", True)
    retentionDays = GetSetting("FormirovanieLetters", "Backup", "RetentionDays", 7)
    lastBackup = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))

    GetBackupSettings = "BACKUP SETTINGS:" & vbCrLf & vbCrLf & _
                        "Automatic backup: " & IIf(autoEnabled, "Enabled", "Disabled") & vbCrLf & _
                        "Retention period: " & retentionDays & " days" & vbCrLf & _
                        "Last backup: " & IIf(lastBackup = DateSerial(1900, 1, 1), "Never", Format$(lastBackup, "dd.mm.yyyy")) & vbCrLf & _
                        "Backup folder: " & GetBackupFolderPath()

    On Error GoTo 0
End Function
