Attribute VB_Name = "ModuleRibbon"
' ======================================================================
' Module: ModuleRibbon
' Author: CreateLetter contributors
' Purpose: Excel Ribbon callbacks and user-configurable folder settings
' Version: 1.0.2 - 29.03.2026
' ======================================================================

Option Explicit

Private Const RibbonSettingsAppName As String = "CreateLetter"
Private Const RibbonSettingsSection As String = "RibbonPaths"
Private Const RibbonSettingTemplateFolder As String = "TemplateFolder"
Private Const RibbonSettingOutputFolder As String = "OutputFolder"
Private Const msoFileDialogFolderPicker As Long = 4

Private ribbonUiHandle As IRibbonUI

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribbonUiHandle = ribbon
End Sub

Public Sub RibbonOpenLetterForm(control As IRibbonControl)
    StartFormirovanieLetters
End Sub

Public Sub RibbonSelectTemplateFolder(control As IRibbonControl)
    PromptAndSaveFolderSetting RibbonSettingTemplateFolder, _
                               t("ribbon.dialog.template_folder", "Select templates folder"), _
                               t("ribbon.msg.template_folder_saved", "Template folder saved:")
End Sub

Public Sub RibbonSelectOutputFolder(control As IRibbonControl)
    PromptAndSaveFolderSetting RibbonSettingOutputFolder, _
                               t("ribbon.dialog.output_folder", "Select output folder"), _
                               t("ribbon.msg.output_folder_saved", "Output folder saved:")
End Sub

Public Sub RibbonShowAbout(control As IRibbonControl)
    MsgBox BuildAboutMessage(), vbInformation, t("ribbon.about.title", "About")
End Sub

Public Function GetConfiguredTemplateFolderPath() As String
    GetConfiguredTemplateFolderPath = GetConfiguredFolderPath(RibbonSettingTemplateFolder)
End Function

Public Function GetConfiguredOutputFolderPath() As String
    GetConfiguredOutputFolderPath = GetConfiguredFolderPath(RibbonSettingOutputFolder)
End Function

Private Function GetConfiguredFolderPath(settingKey As String) As String
    Dim storedPath As String
    storedPath = NormalizeFolderPath(GetSetting(RibbonSettingsAppName, RibbonSettingsSection, settingKey, ""))

    If Len(storedPath) = 0 Then
        GetConfiguredFolderPath = GetDefaultWorkbookFolderPath()
        Exit Function
    End If

    If Dir$(storedPath, vbDirectory) = "" Then
        GetConfiguredFolderPath = GetDefaultWorkbookFolderPath()
        Debug.Print t("ribbon.msg.folder_unavailable", "Configured path is unavailable, using workbook path:") & " " & storedPath
        Exit Function
    End If

    GetConfiguredFolderPath = storedPath
End Function

Private Sub PromptAndSaveFolderSetting(settingKey As String, dialogTitle As String, successMessage As String)
    On Error GoTo DialogError

    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)

    With dialog
        .Title = dialogTitle
        .AllowMultiSelect = False
        .InitialFileName = GetConfiguredFolderPath(settingKey) & "\"

        If .Show = -1 Then
            Dim selectedFolder As String
            selectedFolder = NormalizeFolderPath(.SelectedItems(1))

            SaveSetting RibbonSettingsAppName, RibbonSettingsSection, settingKey, selectedFolder
            MsgBox successMessage & vbCrLf & selectedFolder, vbInformation
        End If
    End With
    Exit Sub

DialogError:
    MsgBox t("ribbon.msg.folder_select_error", "Folder selection error: ") & Err.Description, vbExclamation
End Sub

Private Function BuildAboutMessage() As String
    BuildAboutMessage = t("ribbon.about.name", "CreateLetter") & vbCrLf & vbCrLf & _
                        t("ribbon.about.templates_folder", "Templates folder: ") & GetConfiguredTemplateFolderPath() & vbCrLf & _
                        t("ribbon.about.output_folder", "Output folder: ") & GetConfiguredOutputFolderPath() & vbCrLf & vbCrLf & _
                        t("ribbon.about.open_form_hint", "Use the Excel ribbon to open the form and configure folders.")
End Function

Private Function NormalizeFolderPath(folderPath As String) As String
    NormalizeFolderPath = Trim$(folderPath)

    If Right$(NormalizeFolderPath, 1) = "\" Then
        NormalizeFolderPath = Left$(NormalizeFolderPath, Len(NormalizeFolderPath) - 1)
    End If
End Function

Private Function GetDefaultWorkbookFolderPath() As String
    If Len(Trim$(ThisWorkbook.Path)) > 0 Then
        GetDefaultWorkbookFolderPath = ThisWorkbook.Path
    Else
        GetDefaultWorkbookFolderPath = CurDir$
    End If
End Function
