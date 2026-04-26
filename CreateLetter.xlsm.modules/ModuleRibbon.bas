Attribute VB_Name = "ModuleRibbon"

' ======================================================================

' Module: ModuleRibbon

' Author: CreateLetter contributors

' Purpose: Excel Ribbon callbacks and user-configurable folder settings

' Version: 1.0.7 - 26.04.2026

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



Public Sub RibbonOpenHistoryForm(control As IRibbonControl)

    ShowLetterHistoryModeless

End Sub



Public Sub RibbonOpenMailDispatch(control As IRibbonControl)

    Load frmMailDispatch
    frmMailDispatch.Show vbModeless

End Sub



Public Sub RibbonBuildDispatchRegistry(control As IRibbonControl)

    On Error GoTo RegistryError

    Dim builtCount As Long
    builtCount = BuildDispatchRegistryFromDispatchItems()

    If builtCount > 0 Then
        MsgBox t("dispatch.registry.msg.built", "Registry built from dispatch items.") & vbCrLf & builtCount, _
               vbInformation, _
               t("dispatch.registry.title", "Dispatch registry")
    Else
        MsgBox t("dispatch.registry.msg.no_items", "There are no dispatch items to include in the registry."), _
               vbExclamation, _
               t("dispatch.registry.title", "Dispatch registry")
    End If

    Exit Sub

RegistryError:
    MsgBox t("dispatch.registry.msg.error", "Failed to build the internal dispatch registry: ") & Err.description, _
           vbCritical, _
           t("dispatch.registry.title", "Dispatch registry")
End Sub



Public Sub RibbonSelectTemplateFolder(control As IRibbonControl)

    PromptAndSaveFolderSetting RibbonSettingTemplateFolder, _
                               t("ribbon.dialog.template_folder", GetRibbonTemplateFolderDialogText()), _
                               t("ribbon.msg.template_folder_saved", GetRibbonTemplateFolderSavedText())

End Sub



Public Sub RibbonSelectOutputFolder(control As IRibbonControl)

    PromptAndSaveFolderSetting RibbonSettingOutputFolder, _
                               t("ribbon.dialog.output_folder", GetRibbonOutputFolderDialogText()), _
                               t("ribbon.msg.output_folder_saved", GetRibbonOutputFolderSavedText())

End Sub



Public Sub RibbonShowAbout(control As IRibbonControl)

    MsgBox BuildAboutMessage(), vbInformation, t("ribbon.about.title", GetRibbonAboutTitleText())

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



    If dir$(storedPath, vbDirectory) = "" Then

        GetConfiguredFolderPath = GetDefaultWorkbookFolderPath()

        Debug.Print t("ribbon.msg.folder_unavailable", GetRibbonFolderUnavailableText()) & " " & storedPath

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

    MsgBox t("ribbon.msg.folder_select_error", GetRibbonFolderSelectErrorText()) & Err.description, vbExclamation

End Sub



Private Function BuildAboutMessage() As String

    BuildAboutMessage = t("ribbon.about.name", "CreateLetter") & vbCrLf & vbCrLf & _
                        t("ribbon.about.templates_folder", GetRibbonAboutTemplatesFolderText()) & GetConfiguredTemplateFolderPath() & vbCrLf & _
                        t("ribbon.about.output_folder", GetRibbonAboutOutputFolderText()) & GetConfiguredOutputFolderPath() & vbCrLf & vbCrLf & _
                        t("ribbon.about.open_form_hint", GetRibbonAboutHintText())

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

Private Function GetRibbonTemplateFolderDialogText() As String

    GetRibbonTemplateFolderDialogText = BuildUnicodeText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1091, 32, 1096, 1072, 1073, 1083, 1086, 1085, 1086, 1074)

End Function

Private Function GetRibbonTemplateFolderSavedText() As String

    GetRibbonTemplateFolderSavedText = BuildUnicodeText(1055, 1072, 1087, 1082, 1072, 32, 1096, 1072, 1073, 1083, 1086, 1085, 1086, 1074, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1072, 58)

End Function

Private Function GetRibbonOutputFolderDialogText() As String

    GetRibbonOutputFolderDialogText = BuildUnicodeText(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1087, 1072, 1087, 1082, 1091, 32, 1087, 1080, 1089, 1077, 1084)

End Function

Private Function GetRibbonOutputFolderSavedText() As String

    GetRibbonOutputFolderSavedText = BuildUnicodeText(1055, 1072, 1087, 1082, 1072, 32, 1087, 1080, 1089, 1077, 1084, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1072, 58)

End Function

Private Function GetRibbonAboutTitleText() As String

    GetRibbonAboutTitleText = BuildUnicodeText(1054, 32, 1087, 1088, 1086, 1075, 1088, 1072, 1084, 1084, 1077)

End Function

Private Function GetRibbonFolderUnavailableText() As String

    GetRibbonFolderUnavailableText = BuildUnicodeText(1053, 1072, 1089, 1090, 1088, 1086, 1077, 1085, 1085, 1099, 1081, 32, 1087, 1091, 1090, 1100, 32, 1085, 1077, 1076, 1086, 1089, 1090, 1091, 1087, 1077, 1085, 44, 32, 1080, 1089, 1087, 1086, 1083, 1100, 1079, 1091, 1077, 1090, 1089, 1103, 32, 1087, 1072, 1087, 1082, 1072, 32, 1082, 1085, 1080, 1075, 1080, 58)

End Function

Private Function GetRibbonFolderSelectErrorText() As String

    GetRibbonFolderSelectErrorText = BuildUnicodeText(1054, 1096, 1080, 1073, 1082, 1072, 32, 1074, 1099, 1073, 1086, 1088, 1072, 32, 1087, 1072, 1087, 1082, 1080, 58, 32)

End Function

Private Function GetRibbonAboutTemplatesFolderText() As String

    GetRibbonAboutTemplatesFolderText = BuildUnicodeText(1055, 1072, 1087, 1082, 1072, 32, 1096, 1072, 1073, 1083, 1086, 1085, 1086, 1074, 58, 32)

End Function

Private Function GetRibbonAboutOutputFolderText() As String

    GetRibbonAboutOutputFolderText = BuildUnicodeText(1055, 1072, 1087, 1082, 1072, 32, 1087, 1080, 1089, 1077, 1084, 58, 32)

End Function

Private Function GetRibbonAboutHintText() As String

    GetRibbonAboutHintText = BuildUnicodeText(1048, 1089, 1087, 1086, 1083, 1100, 1079, 1091, 1081, 1090, 1077, 32, 1083, 1077, 1085, 1090, 1091, 32, 69, 120, 99, 101, 108, 44, 32, 1095, 1090, 1086, 1073, 1099, 32, 1086, 1090, 1082, 1088, 1099, 1090, 1100, 32, 1092, 1086, 1088, 1084, 1091, 32, 1080, 32, 1085, 1072, 1089, 1090, 1088, 1086, 1080, 1090, 1100, 32, 1088, 1072, 1073, 1086, 1095, 1080, 1077, 32, 1087, 1072, 1087, 1082, 1080, 46)

End Function

Private Function BuildUnicodeText(ParamArray codePoints() As Variant) As String

    Dim i As Long

    BuildUnicodeText = ""

    For i = LBound(codePoints) To UBound(codePoints)
        BuildUnicodeText = BuildUnicodeText & ChrW(CLng(codePoints(i)))
    Next i

End Function
