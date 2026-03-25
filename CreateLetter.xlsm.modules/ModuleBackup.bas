Attribute VB_Name = "ModuleBackup"
' ======================================================================
' Module: ModuleBackup
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Data backup system
' Version: 1.4.1 — 07.08.2025
' ======================================================================

Option Explicit

Public Sub CreateBackup()
    On Error GoTo BackupError
    
    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backups\"
    
    ' Create folder if it doesn't exist
    If dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If
    
    Dim backupFileName As String
    backupFileName = "FormirovanieLetters_backup_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".xlsx"
    
    Dim backupPath As String
    backupPath = backupFolder & backupFileName
    
    ' Create a copy of the file
    ThisWorkbook.SaveCopyAs backupPath
    
    ' Clean up old backups
    CleanOldBackups backupFolder, 7  ' Keep for 7 days
    
    Debug.Print "Backup created: " & backupFileName
    MsgBox "Backup created successfully!" & vbCrLf & backupPath, vbInformation
    
    Exit Sub
    
BackupError:
    Debug.Print "Error creating backup: " & Err.description
    MsgBox "Error creating backup: " & Err.description, vbCritical
End Sub

Private Sub CleanOldBackups(backupFolder As String, daysToKeep As Integer)
    On Error Resume Next
    
    Dim fileName As String
    fileName = dir(backupFolder & "FormirovanieLetters_backup_*.xlsx")
    
    Do While fileName <> ""
        Dim filePath As String
        filePath = backupFolder & fileName
        
        Dim fileDate As Date
        fileDate = FileDateTime(filePath)
        
        ' Delete files older than the specified number of days
        If Date - fileDate > daysToKeep Then
            Kill filePath
            Debug.Print "Old backup deleted: " & fileName
        End If
        
        fileName = dir
    Loop
    
    On Error GoTo 0
End Sub

Public Sub AutoBackupOnStartup()
    ' Automatic backup on startup (if the last one is older than 24 hours)
    On Error Resume Next
    
    Dim lastBackupDate As Date
    lastBackupDate = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))
    
    If Date - lastBackupDate >= 1 Then  ' More than a day has passed
        CreateBackup
        SaveSetting "FormirovanieLetters", "Backup", "LastBackupDate", Date
    End If
    
    On Error GoTo 0
End Sub

Public Sub ShowBackupInfo()
    ' Show information about backups
    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backups\"
    
    If dir(backupFolder, vbDirectory) = "" Then
        MsgBox "Backup folder not found.", vbInformation
        Exit Sub
    End If
    
    Dim fileName As String
    Dim backupList As String
    Dim backupCount As Integer
    
    fileName = dir(backupFolder & "FormirovanieLetters_backup_*.xlsx")
    backupList = "BACKUP LIST:" & vbCrLf & vbCrLf
    
    Do While fileName <> ""
        Dim filePath As String
        filePath = backupFolder & fileName
        
        Dim fileDate As Date
        fileDate = FileDateTime(filePath)
        
        Dim fileSize As Long
        fileSize = FileLen(filePath)
        
        backupList = backupList & fileName & vbCrLf
        backupList = backupList & "  Date: " & Format(fileDate, "dd.mm.yyyy hh:mm") & vbCrLf
        backupList = backupList & "  Size: " & Format(fileSize \ 1024, "#,##0") & " KB" & vbCrLf & vbCrLf
        
        backupCount = backupCount + 1
        fileName = dir
    Loop
    
    If backupCount = 0 Then
        MsgBox "No backups found.", vbInformation
    Else
        backupList = "Backups found: " & backupCount & vbCrLf & vbCrLf & backupList
        MsgBox backupList, vbInformation, "Backup Information"
    End If
End Sub

Public Sub RestoreFromBackup()
    ' Restore from backup (stub)
    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backups\"
    
    If dir(backupFolder, vbDirectory) = "" Then
        MsgBox "Backup folder not found.", vbExclamation
        Exit Sub
    End If
    
    MsgBox "To restore from a backup:" & vbCrLf & vbCrLf & _
           "1. Close the current file" & vbCrLf & _
           "2. Go to the folder: " & backupFolder & vbCrLf & _
           "3. Select the desired copy and rename it" & vbCrLf & _
           "4. Open the restored file", vbInformation, "Restore from backup"
End Sub

Public Function GetBackupFolderPath() As String
    ' Get the path to the backup folder
    GetBackupFolderPath = ThisWorkbook.Path & "\Backups\"
End Function

Public Function GetLastBackupDate() As Date
    ' Get the date of the last backup
    On Error Resume Next
    GetLastBackupDate = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))
    On Error GoTo 0
End Function

Public Sub SetBackupSettings(enableAutoBackup As Boolean, retentionDays As Integer)
    ' Configure backup settings
    SaveSetting "FormirovanieLetters", "Backup", "AutoBackupEnabled", enableAutoBackup
    SaveSetting "FormirovanieLetters", "Backup", "RetentionDays", retentionDays
    
    MsgBox "Backup settings saved:" & vbCrLf & _
           "Automatic backup: " & IIf(enableAutoBackup, "Enabled", "Disabled") & vbCrLf & _
           "Keep copies for: " & retentionDays & " days", vbInformation
End Sub

Public Function GetBackupSettings() As String
    ' Get current backup settings
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
                       "Last backup: " & IIf(lastBackup = DateSerial(1900, 1, 1), "Never", Format(lastBackup, "dd.mm.yyyy")) & vbCrLf & _
                       "Backup folder: " & GetBackupFolderPath()
    
    On Error GoTo 0
End Function

