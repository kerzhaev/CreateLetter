Attribute VB_Name = "ModuleBackup"

' ======================================================================

' Module: ModuleBackup

' Author: CreateLetter contributors

' Purpose: Workbook backup helpers

' Version: 1.4.6 - 29.03.2026

' ======================================================================



Option Explicit

Private Const BackupFileBaseName As String = "FormirovanieLetters_backup_"

Private Const DefaultBackupRetentionDays As Integer = 7



Public Sub CreateBackup()

    On Error GoTo BackupError



    Dim backupFolder As String

    Dim backupFileName As String

    Dim backupPath As String



    backupFolder = ThisWorkbook.Path & "\Backups\"

    EnsureBackupFolderExists backupFolder



    backupFileName = BuildBackupFileName()

    backupPath = backupFolder & backupFileName



    ThisWorkbook.SaveCopyAs backupPath

    CleanOldBackups backupFolder, GetConfiguredBackupRetentionDays()



    Debug.Print "Backup created: " & backupFileName

    MsgBox t("backup.msg.created_success", "Backup created successfully!") & vbCrLf & backupPath, vbInformation

    Exit Sub



BackupError:

    Debug.Print "Error creating backup: " & Err.description

    MsgBox t("backup.msg.create_error", "Error creating backup: ") & Err.description, vbCritical

End Sub



Private Sub EnsureBackupFolderExists(backupFolder As String)

    On Error GoTo FolderError



    If dir$(backupFolder, vbDirectory) = "" Then

        MkDir backupFolder

    End If

    Exit Sub



FolderError:

    Err.Raise Err.Number, "EnsureBackupFolderExists", Err.description

End Sub



Private Sub CleanOldBackups(backupFolder As String, daysToKeep As Integer)

    On Error GoTo CleanError



    Dim fileName As String

    Dim filePath As String

    Dim fileDate As Date



    fileName = dir$(backupFolder & BackupFileBaseName & "*.xls*")

    Do While fileName <> ""

        filePath = backupFolder & fileName

        fileDate = FileDateTime(filePath)



        If Date - fileDate > daysToKeep Then

            DeleteBackupFile filePath, fileName

        End If



        fileName = dir$

    Loop

    Exit Sub



CleanError:

    Debug.Print "Error cleaning old backups: " & Err.description

End Sub



Private Sub DeleteBackupFile(filePath As String, fileName As String)

    On Error GoTo DeleteError



    Kill filePath

    Debug.Print "Old backup deleted: " & fileName

    Exit Sub



DeleteError:

    Debug.Print "Failed to delete backup '" & fileName & "': " & Err.description

End Sub



Public Sub AutoBackupOnStartup()

    On Error GoTo AutoBackupError



    Dim lastBackupDate As Date

    Dim autoBackupEnabled As Boolean

    lastBackupDate = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))

    autoBackupEnabled = CBool(GetSetting("FormirovanieLetters", "Backup", "AutoBackupEnabled", True))



    If autoBackupEnabled And Date - lastBackupDate >= 1 Then

        CreateBackup

        SaveSetting "FormirovanieLetters", "Backup", "LastBackupDate", Date

    End If

    Exit Sub



AutoBackupError:

    Debug.Print "AutoBackupOnStartup error: " & Err.description

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

    If dir$(backupFolder, vbDirectory) = "" Then

        MsgBox t("backup.msg.folder_not_found", "Backup folder not found."), vbInformation

        Exit Sub

    End If



    fileName = dir$(backupFolder & BackupFileBaseName & "*.xls*")

    backupList = t("backup.msg.list_title", "BACKUP LIST:") & vbCrLf & vbCrLf



    Do While fileName <> ""

        filePath = backupFolder & fileName

        fileDate = FileDateTime(filePath)

        fileSize = FileLen(filePath)



        backupList = backupList & fileName & vbCrLf

        backupList = backupList & t("backup.msg.date_label", "  Date: ") & Format$(fileDate, "dd.mm.yyyy hh:mm") & vbCrLf

        backupList = backupList & t("backup.msg.size_label", "  Size: ") & Format$(fileSize \ 1024, "#,##0") & " KB" & vbCrLf & vbCrLf



        backupCount = backupCount + 1

        fileName = dir$

    Loop



    If backupCount = 0 Then

        MsgBox t("backup.msg.none_found", "No backups found."), vbInformation

    Else

        backupList = t("backup.msg.found_count", "Backups found: ") & backupCount & vbCrLf & vbCrLf & backupList

        MsgBox backupList, vbInformation, t("backup.title.information", "Backup Information")

    End If

End Sub



Public Sub RestoreFromBackup()

    Dim backupFolder As String



    backupFolder = ThisWorkbook.Path & "\Backups\"

    If dir$(backupFolder, vbDirectory) = "" Then

        MsgBox t("backup.msg.folder_not_found", "Backup folder not found."), vbExclamation

        Exit Sub

    End If



    MsgBox t("backup.msg.restore_instructions", "To restore from a backup:") & vbCrLf & vbCrLf & _
           t("backup.msg.restore_step_1", "1. Close the current workbook") & vbCrLf & _
           t("backup.msg.restore_step_2", "2. Open the folder: ") & backupFolder & vbCrLf & _
           t("backup.msg.restore_step_3", "3. Select the desired backup copy") & vbCrLf & _
           t("backup.msg.restore_step_4", "4. Rename and open the restored workbook"), _
           vbInformation, t("backup.title.restore", "Restore from Backup")

End Sub



Public Function GetBackupFolderPath() As String

    GetBackupFolderPath = ThisWorkbook.Path & "\Backups\"

End Function



Public Function GetLastBackupDate() As Date

    On Error GoTo ReadError



    GetLastBackupDate = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))

    Exit Function



ReadError:

    GetLastBackupDate = DateSerial(1900, 1, 1)

    Debug.Print "GetLastBackupDate error: " & Err.description

End Function



Public Sub SetBackupSettings(enableAutoBackup As Boolean, retentionDays As Integer)

    SaveSetting "FormirovanieLetters", "Backup", "AutoBackupEnabled", enableAutoBackup

    SaveSetting "FormirovanieLetters", "Backup", "RetentionDays", retentionDays



    MsgBox t("backup.msg.settings_saved", "Backup settings saved:") & vbCrLf & _
           t("backup.msg.automatic_backup", "Automatic backup: ") & IIf(enableAutoBackup, t("backup.label.enabled", "Enabled"), t("backup.label.disabled", "Disabled")) & vbCrLf & _
           t("backup.msg.keep_copies", "Keep copies for: ") & retentionDays & " days", vbInformation

End Sub



Public Function GetBackupSettings() As String

    On Error GoTo SettingsError



    Dim autoEnabled As Boolean

    Dim retentionDays As Integer

    Dim lastBackup As Date



    autoEnabled = GetSetting("FormirovanieLetters", "Backup", "AutoBackupEnabled", True)

    retentionDays = GetSetting("FormirovanieLetters", "Backup", "RetentionDays", 7)

    lastBackup = GetSetting("FormirovanieLetters", "Backup", "LastBackupDate", DateSerial(1900, 1, 1))



    GetBackupSettings = t("backup.msg.settings_title", "BACKUP SETTINGS:") & vbCrLf & vbCrLf & _
                        t("backup.msg.automatic_backup", "Automatic backup: ") & IIf(autoEnabled, t("backup.label.enabled", "Enabled"), t("backup.label.disabled", "Disabled")) & vbCrLf & _
                        t("backup.msg.retention_period", "Retention period: ") & retentionDays & " days" & vbCrLf & _
                        t("backup.msg.last_backup", "Last backup: ") & IIf(lastBackup = DateSerial(1900, 1, 1), t("backup.label.never", "Never"), Format$(lastBackup, "dd.mm.yyyy")) & vbCrLf & _
                        t("backup.msg.backup_folder", "Backup folder: ") & GetBackupFolderPath()

    Exit Function



SettingsError:

    Debug.Print "GetBackupSettings error: " & Err.description

    GetBackupSettings = t("backup.msg.settings_unavailable", "Backup settings are currently unavailable.")

End Function



Private Function BuildBackupFileName() As String

    BuildBackupFileName = BackupFileBaseName & Format$(Now, "yyyy-mm-dd_hh-mm-ss") & GetWorkbookBackupExtension()

End Function



Private Function GetWorkbookBackupExtension() As String

    Dim workbookName As String

    Dim extensionPosition As Long



    workbookName = ThisWorkbook.Name

    extensionPosition = InStrRev(workbookName, ".")



    If extensionPosition > 0 Then

        GetWorkbookBackupExtension = Mid$(workbookName, extensionPosition)

    Else

        GetWorkbookBackupExtension = ".xlsm"

    End If

End Function



Private Function GetConfiguredBackupRetentionDays() As Integer

    Dim storedRetention As Variant



    storedRetention = GetSetting("FormirovanieLetters", "Backup", "RetentionDays", DefaultBackupRetentionDays)

    If IsNumeric(storedRetention) Then

        GetConfiguredBackupRetentionDays = CInt(storedRetention)

    Else

        GetConfiguredBackupRetentionDays = DefaultBackupRetentionDays

    End If



    If GetConfiguredBackupRetentionDays < 1 Then

        GetConfiguredBackupRetentionDays = DefaultBackupRetentionDays

    End If

End Function



