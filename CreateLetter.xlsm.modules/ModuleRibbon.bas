Attribute VB_Name = "ModuleRibbon"

' ======================================================================

' Module: ModuleRibbon

' Author: CreateLetter contributors

' Purpose: Excel Ribbon callbacks, dispatch actions, and workbook-backed program settings

' Version: 1.4.3 - 03.05.2026

' ======================================================================



Option Explicit



Private Const RibbonSettingsAppName As String = "CreateLetter"

Private Const RibbonSettingsSection As String = "RibbonPaths"

Private Const RibbonSettingTemplateFolder As String = "TemplateFolder"

Private Const RibbonSettingOutputFolder As String = "OutputFolder"

Private Const ProgramSettingsSheetName As String = "ProgramSettings"

Private Const ProgramSettingsTableName As String = "tblProgramSettings"

Private Const ProgramSettingsColumnKey As Long = 1

Private Const ProgramSettingsColumnValue As Long = 2

Private Const ProgramSettingsColumnDescription As Long = 3

Private Const ProgramSettingRequireOutgoingNumber As String = "RequireOutgoingNumber"

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

    OpenMailDispatch

End Sub



Public Sub RibbonOpenDispatchJournal(control As IRibbonControl)

    OpenDispatchJournal

End Sub



Public Sub RibbonReturnDispatchPackage(control As IRibbonControl)

    PromptReturnDispatchPackageToWork

End Sub



Public Sub RibbonCleanupDispatchLegacy(control As IRibbonControl)

    On Error GoTo CleanupError

    Dim legacyCount As Long
    legacyCount = DispatchRepositoryCountLegacyDispatchItems()

    If legacyCount = 0 Then
        MsgBox t("dispatch.cleanup.msg.no_legacy", "Старые строки почтовых отправлений не найдены."), vbInformation, t("dispatch.cleanup.title", "Диагностика отправлений")
        Exit Sub
    End If

    Dim promptText As String
    promptText = t("dispatch.cleanup.confirm", "Найдены старые строки отправлений без статуса и реестра. Удалить их и вернуть связанные письма в работу?") & vbCrLf & CStr(legacyCount)

    If MsgBox(promptText, vbQuestion + vbYesNo, t("dispatch.cleanup.title", "Диагностика отправлений")) <> vbYes Then Exit Sub

    Dim cleanedCount As Long
    cleanedCount = DispatchRepositoryCleanupLegacyDispatchItems()

    MsgBox t("dispatch.cleanup.msg.done", "Очищено старых строк отправлений:") & vbCrLf & CStr(cleanedCount), vbInformation, t("dispatch.cleanup.title", "Диагностика отправлений")
    Exit Sub

CleanupError:

    MsgBox t("dispatch.cleanup.msg.error", "Не удалось очистить старые строки отправлений: ") & Err.description, vbCritical, t("dispatch.cleanup.title", "Диагностика отправлений")

End Sub



Public Sub RibbonBuildDispatchRegistry(control As IRibbonControl)

    On Error GoTo RegistryError

    Dim builtCount As Long
    builtCount = BuildDispatchRegistryFromDispatchItems()

    Dim existingRegistryRows As Long
    existingRegistryRows = CountDispatchRegistryRows()

    If builtCount > 0 Or existingRegistryRows > 0 Then
        Dim printCount As Long
        printCount = BuildPostalRegistryPrintSheet()

        Dim resultMessage As String
        If builtCount > 0 Then
            resultMessage = t("dispatch.registry.msg.built", "Internal registry built from dispatch items.")
            resultMessage = resultMessage & vbCrLf & builtCount
        Else
            resultMessage = t("dispatch.registry.msg.already_current", "There are no new dispatch items. The current registry is already populated and the printable sheet was refreshed.")
            resultMessage = resultMessage & vbCrLf & existingRegistryRows
        End If
        resultMessage = resultMessage & vbCrLf & t("dispatch.registry.msg.print_sheet", "Printable registry sheet built on PostalRegistryPrint.")
        resultMessage = resultMessage & vbCrLf & printCount

        MsgBox resultMessage, vbInformation, t("dispatch.registry.title", "Dispatch registry")
    Else
        MsgBox t("dispatch.registry.msg.no_items", "There are no dispatch items to include in the registry."), vbExclamation, t("dispatch.registry.title", "Dispatch registry")
    End If

    Exit Sub

RegistryError:
    MsgBox t("dispatch.registry.msg.error", "Failed to build the internal dispatch registry: ") & Err.description, vbCritical, t("dispatch.registry.title", "Dispatch registry")
End Sub



Public Sub RibbonConfigurePostalRegistry(control As IRibbonControl)

    ConfigurePostalRegistryPrint

End Sub



Public Sub RibbonShowProgramSettings(control As IRibbonControl)

    ConfigureProgramSettings

End Sub

Public Sub ConfigureProgramSettings()

    On Error GoTo SettingsError

    frmProgramSettings.Show
    Exit Sub

SettingsError:

    MsgBox t("ribbon.settings.msg.error", "Failed to save program settings: ") & Err.description, vbExclamation, t("ribbon.settings.title", "Program settings")

End Sub

Public Function IsOutgoingNumberRequired() As Boolean

    IsOutgoingNumberRequired = IsWorkbookProgramSettingEnabled(ProgramSettingRequireOutgoingNumber, False)

End Function

Public Sub SetOutgoingNumberRequired(required As Boolean)

    Dim storedValue As String
    If required Then
        storedValue = "1"
    Else
        storedValue = "0"
    End If

    SaveWorkbookProgramSetting ProgramSettingRequireOutgoingNumber, storedValue, "Require completed outgoing letter number before leaving the letter step"

End Sub

Private Function IsWorkbookProgramSettingEnabled(settingKey As String, defaultValue As Boolean) As Boolean

    Dim defaultText As String
    If defaultValue Then
        defaultText = "1"
    Else
        defaultText = "0"
    End If

    Dim storedValue As String
    storedValue = LCase$(Trim$(GetWorkbookProgramSetting(settingKey, defaultText)))

    IsWorkbookProgramSettingEnabled = storedValue = "1" Or storedValue = "true" Or storedValue = "yes" Or storedValue = "on"

End Function

Private Function GetWorkbookProgramSetting(settingKey As String, defaultValue As String) As String

    On Error GoTo SettingError

    Dim settingsTable As ListObject
    Set settingsTable = GetProgramSettingsTable()

    Dim settingRow As ListRow
    Set settingRow = FindProgramSettingRow(settingsTable, settingKey)

    If settingRow Is Nothing Then
        GetWorkbookProgramSetting = defaultValue
    Else
        GetWorkbookProgramSetting = CStr(settingRow.Range.Cells(1, ProgramSettingsColumnValue).Value)
    End If

    Exit Function

SettingError:
    Debug.Print "GetWorkbookProgramSetting error: " & Err.description
    GetWorkbookProgramSetting = defaultValue

End Function

Private Sub SaveWorkbookProgramSetting(settingKey As String, settingValue As String, description As String)

    On Error GoTo SaveError

    Dim settingsTable As ListObject
    Set settingsTable = GetProgramSettingsTable()

    Dim settingRow As ListRow
    Set settingRow = FindProgramSettingRow(settingsTable, settingKey)

    If settingRow Is Nothing Then
        Set settingRow = settingsTable.ListRows.Add
    End If

    settingRow.Range.Cells(1, ProgramSettingsColumnKey).Value = settingKey
    settingRow.Range.Cells(1, ProgramSettingsColumnValue).Value = settingValue
    settingRow.Range.Cells(1, ProgramSettingsColumnDescription).Value = description
    settingsTable.Parent.Visible = xlSheetVeryHidden

    Exit Sub

SaveError:
    Err.Raise Err.Number, "SaveWorkbookProgramSetting", Err.description

End Sub

Private Function FindProgramSettingRow(settingsTable As ListObject, settingKey As String) As ListRow

    If settingsTable.DataBodyRange Is Nothing Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 1 To settingsTable.DataBodyRange.Rows.count
        If StrComp(CStr(settingsTable.DataBodyRange.Cells(rowIndex, ProgramSettingsColumnKey).Value), settingKey, vbTextCompare) = 0 Then
            Set FindProgramSettingRow = settingsTable.ListRows(rowIndex)
            Exit Function
        End If
    Next rowIndex

End Function

Private Function GetProgramSettingsTable() As ListObject

    Dim settingsSheet As Worksheet
    Set settingsSheet = GetOrCreateProgramSettingsSheet()

    Dim tableIndex As Long
    For tableIndex = 1 To settingsSheet.ListObjects.count
        If settingsSheet.ListObjects(tableIndex).Name = ProgramSettingsTableName Then
            Set GetProgramSettingsTable = settingsSheet.ListObjects(tableIndex)
            EnsureProgramSettingsSeed GetProgramSettingsTable
            settingsSheet.Visible = xlSheetVeryHidden
            Exit Function
        End If
    Next tableIndex

    PrepareProgramSettingsHeaders settingsSheet

    Dim settingsRange As Range
    Set settingsRange = settingsSheet.Range(settingsSheet.Cells(1, 1), settingsSheet.Cells(2, ProgramSettingsColumnDescription))

    Dim settingsTable As ListObject
    Set settingsTable = settingsSheet.ListObjects.Add(xlSrcRange, settingsRange, , xlYes)
    settingsTable.Name = ProgramSettingsTableName

    Set GetProgramSettingsTable = settingsTable
    EnsureProgramSettingsSeed settingsTable
    settingsSheet.Visible = xlSheetVeryHidden

End Function

Private Function GetOrCreateProgramSettingsSheet() As Worksheet

    On Error Resume Next
    Set GetOrCreateProgramSettingsSheet = ThisWorkbook.Worksheets(ProgramSettingsSheetName)
    On Error GoTo 0

    If GetOrCreateProgramSettingsSheet Is Nothing Then
        Set GetOrCreateProgramSettingsSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        GetOrCreateProgramSettingsSheet.Name = ProgramSettingsSheetName
    End If

    GetOrCreateProgramSettingsSheet.Visible = xlSheetVisible

End Function

Private Sub PrepareProgramSettingsHeaders(settingsSheet As Worksheet)

    settingsSheet.Cells(1, ProgramSettingsColumnKey).Value = "SettingKey"
    settingsSheet.Cells(1, ProgramSettingsColumnValue).Value = "SettingValue"
    settingsSheet.Cells(1, ProgramSettingsColumnDescription).Value = "Description"

    If Len(Trim$(CStr(settingsSheet.Cells(2, ProgramSettingsColumnKey).Value))) = 0 Then
        settingsSheet.Cells(2, ProgramSettingsColumnKey).Value = ProgramSettingRequireOutgoingNumber
        settingsSheet.Cells(2, ProgramSettingsColumnValue).Value = "0"
        settingsSheet.Cells(2, ProgramSettingsColumnDescription).Value = "Require completed outgoing letter number before leaving the letter step"
    End If

End Sub

Private Sub EnsureProgramSettingsSeed(settingsTable As ListObject)

    Dim settingRow As ListRow
    Set settingRow = FindProgramSettingRow(settingsTable, ProgramSettingRequireOutgoingNumber)

    If settingRow Is Nothing Then
        Set settingRow = settingsTable.ListRows.Add
        settingRow.Range.Cells(1, ProgramSettingsColumnKey).Value = ProgramSettingRequireOutgoingNumber
        settingRow.Range.Cells(1, ProgramSettingsColumnValue).Value = "0"
        settingRow.Range.Cells(1, ProgramSettingsColumnDescription).Value = "Require completed outgoing letter number before leaving the letter step"
    End If

End Sub



Public Sub RibbonExportPostalRegistryPdf(control As IRibbonControl)

    On Error GoTo ExportError

    If Not ConfirmPostalRegistryPdfWithUnpackedLetters() Then Exit Sub

    Dim pdfPath As String
    pdfPath = ExportPostalRegistryPrint()

    If Len(Trim$(pdfPath)) > 0 Then
        MsgBox t("postal.registry.pdf.msg.exported", "Postal registry PDF exported:") & vbCrLf & pdfPath, vbInformation, t("postal.registry.pdf.title", "Postal registry PDF")
    Else
        MsgBox t("postal.registry.pdf.msg.not_exported", "Postal registry PDF was not exported."), vbExclamation, t("postal.registry.pdf.title", "Postal registry PDF")
    End If

    Exit Sub

ExportError:
    MsgBox t("postal.registry.pdf.msg.error", "Failed to export postal registry PDF: ") & Err.description, vbCritical, t("postal.registry.pdf.title", "Postal registry PDF")
End Sub

Public Function ConfirmPostalRegistryPdfWithUnpackedLetters() As Boolean
    ConfirmPostalRegistryPdfWithUnpackedLetters = True

    Dim sampleText As String
    Dim unpackedCount As Long
    unpackedCount = DispatchRepositoryCountAvailableUnpackedLetters(sampleText)

    If unpackedCount = 0 Then Exit Function

    Dim promptText As String
    promptText = t("postal.registry.pdf.confirm.unpacked_letters", "There are letters in history that have not been added to dispatch packages. If you continue, they will not be included in the current PDF registry.") & vbCrLf & CStr(unpackedCount)
    If Len(Trim$(sampleText)) > 0 Then promptText = promptText & vbCrLf & vbCrLf & sampleText
    promptText = promptText & vbCrLf & vbCrLf & t("postal.registry.pdf.confirm.continue", "Continue printing the PDF registry?")

    If MsgBox(promptText, vbQuestion + vbYesNo, t("postal.registry.pdf.title", "OPS registry")) <> vbYes Then ConfirmPostalRegistryPdfWithUnpackedLetters = False
End Function



Public Sub RibbonPrepareEnvelopePrint(control As IRibbonControl)

    On Error GoTo PrepareError

    Dim preparedCount As Long
    preparedCount = PrepareEnvelopePrint()

    If preparedCount > 0 Then
        MsgBox t("dispatch.layouts.msg.prepared", "Envelope layouts prepared.") & vbCrLf & preparedCount, _
               vbInformation, _
               t("dispatch.layouts.title", "Envelope layouts")
    Else
        MsgBox t("dispatch.layouts.msg.no_items", "There are no dispatch items to prepare for envelope layouts."), _
               vbExclamation, _
               t("dispatch.layouts.title", "Envelope layouts")
    End If

    Exit Sub

PrepareError:
    MsgBox t("dispatch.layouts.msg.error", "Failed to prepare envelope layouts: ") & Err.description, _
           vbCritical, _
           t("dispatch.layouts.title", "Envelope layouts")
End Sub

Public Sub RibbonPrepareEnvelopePreviewGrid(control As IRibbonControl)

    On Error GoTo PreviewError

    Dim preparedCount As Long
    preparedCount = PrepareEnvelopePreviewGrid()

    If preparedCount > 0 Then
        MsgBox t("dispatch.layouts.preview.msg.prepared", "Envelope preview grid prepared.") & vbCrLf & preparedCount, vbInformation, t("dispatch.layouts.preview.title", "Envelope preview grid")
    Else
        MsgBox t("dispatch.layouts.msg.no_items", "There are no dispatch items to prepare for envelope layouts."), vbExclamation, t("dispatch.layouts.preview.title", "Envelope preview grid")
    End If

    Exit Sub

PreviewError:
    MsgBox t("dispatch.layouts.preview.msg.error", "Failed to prepare envelope preview grid: ") & Err.description, vbCritical, t("dispatch.layouts.preview.title", "Envelope preview grid")
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

    Dim messageText As String
    messageText = t("ribbon.about.name", "CreateLetter")
    messageText = messageText & vbCrLf & vbCrLf
    messageText = messageText & t("ribbon.about.pipeline.title", "Workflow:")
    messageText = messageText & vbCrLf & t("ribbon.about.pipeline.step1", "1. Create a letter")
    messageText = messageText & vbCrLf & t("ribbon.about.pipeline.step2", "2. Add letters to a package")
    messageText = messageText & vbCrLf & t("ribbon.about.pipeline.step3", "3. Build the OPS registry")
    messageText = messageText & vbCrLf & t("ribbon.about.pipeline.step4", "4. Prepare envelopes")
    messageText = messageText & vbCrLf & t("ribbon.about.pipeline.step5", "5. Export the OPS registry")
    messageText = messageText & vbCrLf & vbCrLf
    messageText = messageText & t("ribbon.about.templates_folder", GetRibbonAboutTemplatesFolderText()) & GetConfiguredTemplateFolderPath()
    messageText = messageText & vbCrLf & t("ribbon.about.output_folder", GetRibbonAboutOutputFolderText()) & GetConfiguredOutputFolderPath()
    messageText = messageText & vbCrLf & vbCrLf
    messageText = messageText & t("ribbon.about.version", "Version: ") & CreateLetterApplicationVersion

    BuildAboutMessage = messageText

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
