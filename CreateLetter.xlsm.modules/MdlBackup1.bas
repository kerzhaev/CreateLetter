Attribute VB_Name = "MdlBackup1"

' ======================================================================

' Module: MdlBackup1

' Purpose: Legacy VBA snapshot and workbook snapshot helpers

' Version: 1.0.4 - 29.03.2026

' Notes:

' - This module is kept for compatibility with legacy admin workflows.

' - It is not part of the main end-user runtime path.

' ======================================================================



Option Explicit



Public Sub CreateProjectVbaSnapshot()

    CreateVBASnapshot

End Sub



Public Sub CreateVBASnapshot()

    Dim vbComp As Object

    Dim exportPath As String

    Dim fileName As String

    Dim timeStamp As String

    Dim fso As Object

    Dim infoFile As String

    Dim fileNum As Integer



    timeStamp = Format$(Now, "yyyy-mm-dd_hh-mm-ss")

    exportPath = ThisWorkbook.Path & "\VBA_Snapshots\Snapshot_" & timeStamp & "\"



    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(exportPath) Then

        fso.CreateFolder Left$(exportPath, Len(exportPath) - 1)

    End If



    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        Select Case vbComp.Type

            Case 1: fileName = vbComp.Name & ".bas"

            Case 2: fileName = vbComp.Name & ".cls"

            Case 3: fileName = vbComp.Name & ".frm"

            Case 100: fileName = vbComp.Name & ".cls"

            Case Else: fileName = vbComp.Name & ".txt"

        End Select



        vbComp.Export exportPath & fileName

        Debug.Print "Exported: " & fileName

    Next vbComp



    infoFile = exportPath & "SnapshotInfo.txt"

    fileNum = FreeFile

    Open infoFile For Output As #fileNum

    Print #fileNum, "VBA Project Snapshot"

    Print #fileNum, "Created at: " & Format$(Now, "dd.mm.yyyy hh:mm:ss")

    Print #fileNum, "Workbook: " & ThisWorkbook.Name

    Print #fileNum, "Path: " & ThisWorkbook.FullName

    Print #fileNum, "Components: " & ThisWorkbook.VBProject.VBComponents.count

    Print #fileNum, ""

    Print #fileNum, "Component list:"

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        Print #fileNum, "- " & vbComp.Name & " (Type: " & GetComponentTypeName(vbComp.Type) & ")"

    Next vbComp

    Close #fileNum



    MsgBox t("snapshot.msg.vba_created_success", "VBA snapshot created successfully!") & vbCrLf & _
           t("snapshot.msg.folder_label", "Folder: ") & exportPath & vbCrLf & _
           t("snapshot.msg.exported_components", "Exported components: ") & ThisWorkbook.VBProject.VBComponents.count, _
           vbInformation, t("snapshot.title.vba_snapshot", "VBA Snapshot")



    Shell "explorer.exe " & exportPath, vbNormalFocus

End Sub



Public Function GetComponentTypeName(componentType As Integer) As String

    Select Case componentType

        Case 1: GetComponentTypeName = "Standard Module"

        Case 2: GetComponentTypeName = "Class Module"

        Case 3: GetComponentTypeName = "UserForm"

        Case 100: GetComponentTypeName = "Document Module"

        Case Else: GetComponentTypeName = "Unknown Type"

    End Select

End Function



Public Sub RestoreVbaSnapshot()

    RestoreFromSnapshot

End Sub



Public Sub RestoreFromSnapshot()

    Dim importPath As String

    Dim fso As Object

    Dim folder As Object

    Dim file As Object

    Dim response As VbMsgBoxResult

    Dim componentIndex As Long



    importPath = SelectSnapshotFolder()

    If importPath = "" Then Exit Sub



    Set fso = CreateObject("Scripting.FileSystemObject")

    Set folder = fso.GetFolder(importPath)



    response = MsgBox(t("snapshot.msg.restore_warning", "Warning!" & vbCrLf & _
                      "This operation removes the current VBA modules and restores them from the selected snapshot." & vbCrLf & _
                      "Do you want to continue?"), _
                      vbYesNo + vbExclamation, t("snapshot.title.confirm_restore", "Confirm Restore"))

    If response = vbNo Then Exit Sub



    For componentIndex = ThisWorkbook.VBProject.VBComponents.count To 1 Step -1

        With ThisWorkbook.VBProject.VBComponents(componentIndex)

            If .Type = 1 Or .Type = 2 Or .Type = 3 Then

                ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(componentIndex)

            End If

        End With

    Next componentIndex



    For Each file In folder.Files

        If Right$(LCase$(file.Name), 4) = ".bas" Or _
           Right$(LCase$(file.Name), 4) = ".cls" Or _
           Right$(LCase$(file.Name), 4) = ".frm" Then

            ThisWorkbook.VBProject.VBComponents.Import file.Path

            Debug.Print "Imported: " & file.Name

        End If

    Next file



    MsgBox t("snapshot.msg.restore_complete", "Snapshot restored successfully!") & vbCrLf & _
           t("snapshot.msg.folder_label", "Folder: ") & importPath, _
           vbInformation, t("snapshot.title.restore_complete", "Restore Complete")

End Sub



Public Function SelectVbaSnapshotFolder() As String

    SelectVbaSnapshotFolder = SelectSnapshotFolder()

End Function



Public Function SelectSnapshotFolder() As String

    Dim snapshotsPath As String

    Dim selectedPath As String



    snapshotsPath = ThisWorkbook.Path & "\VBA_Snapshots\"



    If dir$(snapshotsPath, vbDirectory) = "" Then

        MsgBox t("snapshot.msg.folder_not_found", "Snapshots folder not found: ") & snapshotsPath, vbExclamation

        SelectSnapshotFolder = ""

        Exit Function

    End If



    selectedPath = InputBox(t("snapshot.prompt.select_folder", "Enter the snapshot folder name:" & vbCrLf & _
                            "Available root: ") & snapshotsPath, _
                            t("snapshot.title.select_folder", "Select Snapshot"), t("snapshot.prompt.default_folder", "Snapshot_"))



    If selectedPath <> "" Then

        SelectSnapshotFolder = snapshotsPath & selectedPath & "\"

    Else

        SelectSnapshotFolder = ""

    End If

End Function



Public Sub TagAllModulesWithSnapshotVersion()

    AddVersionTagsToAllModules

End Sub



Public Sub AddVersionTagsToAllModules()

    Dim vbComp As Object

    Dim codeModule As Object

    Dim versionTag As String

    Dim currentDate As String



    currentDate = Format$(Now, "dd.mm.yyyy hh:mm")

    versionTag = "' === Snapshot Tag === " & currentDate & " ==="



    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        If vbComp.Type = 1 Then

            Set codeModule = vbComp.codeModule

            codeModule.InsertLines 1, versionTag

            codeModule.InsertLines 2, "' Snapshot tag inserted: " & currentDate

            codeModule.InsertLines 3, ""

            Debug.Print "Version tag inserted into module: " & vbComp.Name

        End If

    Next vbComp



    MsgBox t("snapshot.msg.version_tags_inserted", "Version tags inserted into all standard modules."), vbInformation

End Sub



Public Sub CreateWorkbookFileSnapshot()

    CreateWorkbookSnapshot

End Sub



Public Sub CreateWorkbookSnapshot()

    Dim originalPath As String

    Dim snapshotPath As String

    Dim timeStamp As String

    Dim description As String

    Dim fileName As String



    description = InputBox(t("snapshot.prompt.workbook_label", "Enter a short snapshot label:"), _
                           t("snapshot.title.workbook_snapshot", "Workbook Snapshot"), t("snapshot.prompt.workbook_default", "manual_snapshot"))

    If description = "" Then Exit Sub



    timeStamp = Format$(Now, "yyyy-mm-dd_hh-mm")

    fileName = Replace$(ThisWorkbook.Name, ".xlsm", "") & "_" & description & "_" & timeStamp & ".xlsm"



    originalPath = ThisWorkbook.Path

    snapshotPath = originalPath & "\Snapshots\"



    If dir$(snapshotPath, vbDirectory) = "" Then

        MkDir snapshotPath

    End If



    ThisWorkbook.SaveCopyAs snapshotPath & fileName



    MsgBox t("snapshot.msg.workbook_created", "Workbook snapshot created!") & vbCrLf & _
           t("snapshot.msg.file_label", "File: ") & fileName & vbCrLf & _
           t("snapshot.msg.folder_label", "Folder: ") & snapshotPath, _
           vbInformation, t("snapshot.title.workbook_snapshot", "Workbook Snapshot")



    Shell "explorer.exe " & snapshotPath, vbNormalFocus

End Sub



Public Sub CreateFullSnapshotBundle()

    QuickSnapshot

End Sub



Public Sub QuickSnapshot()

    AddVersionTagsToAllModules

    CreateVBASnapshot

    CreateWorkbookSnapshot

End Sub



Public Sub OpenLetterCreatorForm()

    ModuleMain.ShowLetterCreator

End Sub



Public Sub ShowLetterCreatorLegacy()

    frmLetterCreator.Show vbModeless

End Sub



