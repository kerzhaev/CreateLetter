VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMailDispatch 

   Caption         =   "Mail dispatch v1.0.0"

   ClientHeight    =   4680

   ClientLeft      =   120

   ClientTop       =   465

   ClientWidth     =   6480

   OleObjectBlob   =   "frmMailDispatch.frx":0000

   StartUpPosition =   1  'CenterOwner

End

Attribute VB_Name = "frmMailDispatch"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False





' ======================================================================

' Form: frmMailDispatch v1.4.4

' Author: CreateLetter contributors

' Date: 04.05.2026

' Purpose: Thin-shell UI for preparing RPBS-aware dispatch packages from existing letters with small-screen fallback

' ======================================================================

Option Explicit



Private allAvailableLettersData As Collection

Private availableLettersData As Collection

Private packageLettersData As Collection

Private envelopeFormats As Collection

Private senderItems As Collection

Private dynamicButtonHandlers As Collection

Private lblDispatchSearch As MSForms.Label

Private txtDispatchSearch As MSForms.TextBox

Private lblDispatchPackage As MSForms.Label

Private lstDispatchPackage As MSForms.ListBox

Private btnDispatchAddToPackage As MSForms.CommandButton

Private btnDispatchRemoveFromPackage As MSForms.CommandButton

Private btnDispatchFinalizeRegistry As MSForms.CommandButton

Private lblDispatchRegistryNumber As MSForms.Label

Private txtDispatchRegistryNumber As MSForms.TextBox

Private lblDispatchRegistryDate As MSForms.Label

Private txtDispatchRegistryDate As MSForms.TextBox

Private cmbDispatchMailType As MSForms.ComboBox

Private pendingDoubleClickAction As String

Private pendingDoubleClickIndex As Long

Private doubleClickActionScheduled As Boolean

Private pendingDoubleClickRunAt As Date

Private workingRegistryPromptShown As Boolean



Private Sub UserForm_Initialize()

    Set allAvailableLettersData = New Collection

    Set availableLettersData = New Collection

    Set packageLettersData = New Collection

    Set envelopeFormats = New Collection

    Set senderItems = New Collection

    Set dynamicButtonHandlers = New Collection

    pendingDoubleClickIndex = -1
    pendingDoubleClickRunAt = 0
    workingRegistryPromptShown = False
    RegisterActiveMailDispatchForm Me



    EnsureDynamicControls

    ApplyFormSettings

    ApplyResponsiveLayout

    ApplyLocalizedCaptions

    ConfigureLists

    LoadDispatchData

    SelectDefaultValues

    UpdateDispatchPreview

End Sub



Private Sub UserForm_Terminate()

    On Error Resume Next

    CancelPendingDoubleClickSchedule
    UnregisterActiveMailDispatchForm Me
    On Error GoTo 0

End Sub



Private Sub ApplyFormSettings()

    With Me

        .Caption = t("form.mail_dispatch.title", "Mail dispatch") & " v" & CreateLetterApplicationVersion

        .backColor = RGB(248, 248, 248)

    End With

End Sub



Private Sub ApplyResponsiveLayout()

    Const FORM_WIDTH As Single = 920
    Const FORM_HEIGHT As Single = 590
    Const LEFT_COLUMN_LEFT As Single = 12
    Const LIST_WIDTH As Single = 376
    Const MIDDLE_BUTTON_LEFT As Single = 400
    Const MIDDLE_BUTTON_WIDTH As Single = 70
    Const PACKAGE_COLUMN_LEFT As Single = 482
    Const CONTENT_TOP As Single = 18
    Const SEARCH_TOP As Single = 44
    Const LIST_TOP As Single = 94
    Const LIST_HEIGHT As Single = 230
    Const BUTTON_TOP As Single = 176
    Const PARAM_TOP As Single = 346
    Const ROW_HEIGHT As Single = 25
    Const PREVIEW_TOP As Single = 350
    Const COMMENT_TOP As Single = 450
    Const RIGHT_ACTION_TOP As Single = 445
    Const LEFT_ACTION_TOP As Single = 532



    Me.StartUpPosition = 1

    FitFormToVisibleScreen FORM_WIDTH, FORM_HEIGHT



    lblDispatchLetters.Left = LEFT_COLUMN_LEFT

    lblDispatchLetters.Top = CONTENT_TOP

    lblDispatchLetters.Width = LIST_WIDTH



    lblDispatchSearch.Left = LEFT_COLUMN_LEFT

    lblDispatchSearch.Top = SEARCH_TOP

    lblDispatchSearch.Width = LIST_WIDTH



    txtDispatchSearch.Left = LEFT_COLUMN_LEFT

    txtDispatchSearch.Top = SEARCH_TOP + 20

    txtDispatchSearch.Width = LIST_WIDTH

    txtDispatchSearch.Height = 22



    lstDispatchLetters.Left = LEFT_COLUMN_LEFT

    lstDispatchLetters.Top = LIST_TOP

    lstDispatchLetters.Width = LIST_WIDTH

    lstDispatchLetters.Height = LIST_HEIGHT



    btnDispatchRefresh.Left = LEFT_COLUMN_LEFT

    btnDispatchRefresh.Top = LEFT_ACTION_TOP

    btnDispatchRefresh.Width = 118



    lblDispatchPackage.Left = PACKAGE_COLUMN_LEFT

    lblDispatchPackage.Top = CONTENT_TOP

    lblDispatchPackage.Width = LIST_WIDTH



    lstDispatchPackage.Left = PACKAGE_COLUMN_LEFT

    lstDispatchPackage.Top = LIST_TOP

    lstDispatchPackage.Width = LIST_WIDTH

    lstDispatchPackage.Height = LIST_HEIGHT



    btnDispatchAddToPackage.Left = MIDDLE_BUTTON_LEFT

    btnDispatchAddToPackage.Top = BUTTON_TOP

    btnDispatchAddToPackage.Width = MIDDLE_BUTTON_WIDTH

    btnDispatchAddToPackage.Height = 30



    btnDispatchRemoveFromPackage.Left = MIDDLE_BUTTON_LEFT

    btnDispatchRemoveFromPackage.Top = BUTTON_TOP + 42

    btnDispatchRemoveFromPackage.Width = MIDDLE_BUTTON_WIDTH

    btnDispatchRemoveFromPackage.Height = 30



    lblDispatchSender.Left = PACKAGE_COLUMN_LEFT

    lblDispatchSender.Top = PARAM_TOP

    lblDispatchSender.Width = 70

    cmbDispatchSender.Left = PACKAGE_COLUMN_LEFT + 78

    cmbDispatchSender.Top = PARAM_TOP - 3

    cmbDispatchSender.Width = 190


    lblDispatchEnvelopeFormat.Left = PACKAGE_COLUMN_LEFT

    lblDispatchEnvelopeFormat.Top = PARAM_TOP + ROW_HEIGHT

    lblDispatchEnvelopeFormat.Width = 50

    cmbDispatchEnvelopeFormat.Left = PACKAGE_COLUMN_LEFT + 58

    cmbDispatchEnvelopeFormat.Top = PARAM_TOP + ROW_HEIGHT - 3

    cmbDispatchEnvelopeFormat.Width = 68


    lblDispatchMailType.Left = PACKAGE_COLUMN_LEFT + 146

    lblDispatchMailType.Top = PARAM_TOP + ROW_HEIGHT

    txtDispatchMailType.Visible = False

    lblDispatchMailType.Width = 35

    cmbDispatchMailType.Left = PACKAGE_COLUMN_LEFT + 188

    cmbDispatchMailType.Top = PARAM_TOP + ROW_HEIGHT - 3

    cmbDispatchMailType.Width = 182


    lblDispatchRegistryNumber.Left = PACKAGE_COLUMN_LEFT

    lblDispatchRegistryNumber.Top = PARAM_TOP + (ROW_HEIGHT * 2)

    lblDispatchRegistryNumber.Width = 84

    txtDispatchRegistryNumber.Left = PACKAGE_COLUMN_LEFT + 90

    txtDispatchRegistryNumber.Top = PARAM_TOP + (ROW_HEIGHT * 2) - 3

    txtDispatchRegistryNumber.Width = 88


    lblDispatchRegistryDate.Left = PACKAGE_COLUMN_LEFT + 196

    lblDispatchRegistryDate.Top = PARAM_TOP + (ROW_HEIGHT * 2)

    lblDispatchRegistryDate.Width = 48

    txtDispatchRegistryDate.Left = PACKAGE_COLUMN_LEFT + 250

    txtDispatchRegistryDate.Top = PARAM_TOP + (ROW_HEIGHT * 2) - 3

    txtDispatchRegistryDate.Width = 120


    lblDispatchMass.Visible = False

    txtDispatchMass.Visible = False

    lblDispatchDeclaredValue.Visible = False

    txtDispatchDeclaredValue.Visible = False



    lblDispatchPreview.Left = LEFT_COLUMN_LEFT

    lblDispatchPreview.Top = PREVIEW_TOP - 22

    lblDispatchPreview.Width = LIST_WIDTH

    txtDispatchPreview.Left = LEFT_COLUMN_LEFT

    txtDispatchPreview.Top = PREVIEW_TOP

    txtDispatchPreview.Width = LIST_WIDTH

    txtDispatchPreview.Height = 76



    lblDispatchComment.Left = LEFT_COLUMN_LEFT

    lblDispatchComment.Top = COMMENT_TOP - 22

    lblDispatchComment.Width = LIST_WIDTH

    txtDispatchComment.Left = LEFT_COLUMN_LEFT

    txtDispatchComment.Top = COMMENT_TOP

    txtDispatchComment.Width = LIST_WIDTH

    txtDispatchComment.Height = 56



    btnDispatchCreate.Left = PACKAGE_COLUMN_LEFT

    btnDispatchCreate.Top = RIGHT_ACTION_TOP

    btnDispatchCreate.Width = 162



    btnDispatchFinalizeRegistry.Left = PACKAGE_COLUMN_LEFT + 178

    btnDispatchFinalizeRegistry.Top = RIGHT_ACTION_TOP

    btnDispatchFinalizeRegistry.Width = 190

    btnDispatchFinalizeRegistry.Height = 28



    btnDispatchClose.Left = 236

    btnDispatchClose.Top = LEFT_ACTION_TOP

    btnDispatchClose.Width = 120

    btnDispatchClose.Height = 28

End Sub

Private Sub FitFormToVisibleScreen(designWidth As Single, designHeight As Single)

    On Error GoTo ScreenError

    Dim maxWidth As Single
    maxWidth = Application.UsableWidth - 24

    Dim maxHeight As Single
    maxHeight = Application.UsableHeight - 48

    If maxWidth < 420 Then maxWidth = 420
    If maxHeight < 360 Then maxHeight = 360

    Me.Width = designWidth
    Me.Height = designHeight
    Me.ScrollLeft = 0
    Me.ScrollTop = 0

    If designWidth > maxWidth Or designHeight > maxHeight Then
        If designWidth > maxWidth Then Me.Width = maxWidth
        If designHeight > maxHeight Then Me.Height = maxHeight
        Me.ScrollBars = fmScrollBarsBoth
        Me.ScrollWidth = designWidth
        Me.ScrollHeight = designHeight
    Else
        Me.ScrollBars = fmScrollBarsNone
        Me.ScrollWidth = designWidth
        Me.ScrollHeight = designHeight
    End If

    Exit Sub

ScreenError:

    Me.Width = designWidth
    Me.Height = designHeight
    Me.ScrollBars = fmScrollBarsVertical
    Me.ScrollHeight = designHeight

End Sub



Private Sub ApplyLocalizedCaptions()

    SetLocalizedCaption "lblDispatchLetters", "form.mail_dispatch.label.available_letters", "Awaiting dispatch"

    lblDispatchSearch.Caption = t("form.mail_dispatch.label.search_letters", "Search letters")

    SetLocalizedCaption "lblDispatchSender", "form.mail_dispatch.label.sender", "Sender"

    SetLocalizedCaption "lblDispatchEnvelopeFormat", "form.mail_dispatch.label.envelope_format_short", "Format"

    SetLocalizedCaption "lblDispatchMailType", "form.mail_dispatch.label.mail_type_short", "Type"

    SetLocalizedCaption "lblDispatchRegistryNumber", "form.mail_dispatch.label.registry_number_short", "Registry"

    SetLocalizedCaption "lblDispatchRegistryDate", "form.mail_dispatch.label.registry_date_short", "Date"

    SetLocalizedCaption "lblDispatchComment", "form.mail_dispatch.label.comment", "Comment"

    SetLocalizedCaption "lblDispatchPreview", "form.mail_dispatch.label.preview", "Preview"

    lblDispatchPackage.Caption = t("form.mail_dispatch.label.package_letters", "Envelope package contents")



    btnDispatchRefresh.Caption = t("form.mail_dispatch.button.refresh", "Refresh")

    btnDispatchCreate.Caption = t("form.mail_dispatch.button.create_package", "Prepare envelope")

    btnDispatchClose.Caption = t("form.mail_dispatch.button.close", "Close")

    btnDispatchFinalizeRegistry.Caption = t("form.mail_dispatch.button.finalize_registry", "Finalize batch")

    btnDispatchAddToPackage.Caption = t("form.mail_dispatch.button.add_to_package", "Put in")

    btnDispatchRemoveFromPackage.Caption = t("form.mail_dispatch.button.remove_from_package", "Return")

    btnDispatchAddToPackage.ControlTipText = t("form.mail_dispatch.tip.add_to_package", "Put selected letters into the envelope package")

    btnDispatchRemoveFromPackage.ControlTipText = t("form.mail_dispatch.tip.remove_from_package", "Return selected letters from the package")

    btnDispatchFinalizeRegistry.ControlTipText = t("form.mail_dispatch.tip.finalize_registry", "Export the OPS registry PDF and close the mail batch")



    cmbDispatchMailType.ControlTipText = t("form.mail_dispatch.tip.mail_type", "Select mail type for the envelope mark")

    txtDispatchSearch.ControlTipText = t("form.mail_dispatch.tip.search_letters", "Type number, date, addressee, or letter text to filter the list")

    txtDispatchComment.ControlTipText = t("form.mail_dispatch.tip.comment", "Short operational comment for the dispatch item")

    txtDispatchRegistryNumber.ControlTipText = t("form.mail_dispatch.tip.registry_number", "Internal registry number for this package")

    txtDispatchRegistryDate.ControlTipText = t("form.mail_dispatch.tip.registry_date", "Internal registry date in dd.mm.yyyy format")

End Sub



Private Sub ConfigureLists()

    lstDispatchLetters.MultiSelect = fmMultiSelectMulti

    lstDispatchPackage.MultiSelect = fmMultiSelectMulti

    lstDispatchLetters.Font.Size = 10

    lstDispatchPackage.Font.Size = 10

    txtDispatchPreview.MultiLine = True

    txtDispatchPreview.ScrollBars = fmScrollBarsVertical
    cmbDispatchMailType.Style = fmStyleDropDownList

End Sub



Private Sub LoadDispatchData()

    LoadLettersList

    LoadSendersList

    LoadEnvelopeFormatList
    LoadMailTypeList

End Sub



Private Sub LoadLettersList()

    Dim rawLetters As Collection

    Set rawLetters = RepositoryLoadLetterHistoryData()



    Dim queuedKeys As Object

    Set queuedKeys = DispatchRepositoryGetQueuedLetterKeySet()



    Set allAvailableLettersData = New Collection



    Dim i As Long

    For i = 1 To rawLetters.count

        Dim record As clsLetterHistoryRecord

        Set record = rawLetters(i)



        If IsDispatchRecordAvailable(record, queuedKeys) Then

            allAvailableLettersData.Add record

        End If

    Next i



    ApplyAvailableLettersFilter

    RebindPackageLettersList

End Sub

Private Function IsDispatchRecordAvailable(record As clsLetterHistoryRecord, queuedKeys As Object) As Boolean

    If record Is Nothing Then Exit Function

    If Not queuedKeys Is Nothing Then
        If queuedKeys.Exists(BuildHistoryRecordKey(record)) Then Exit Function
    End If

    Dim packedFlag As String
    packedFlag = UCase$(Trim$(record.DispatchPackedFlag))

    If Len(packedFlag) > 0 Then
        If packedFlag <> UCase$(t("history.dispatch_status.not_packed", "Нет")) And packedFlag <> "NO" Then Exit Function
    End If

    If Len(Trim$(record.DispatchBatchId)) > 0 Then Exit Function
    If Len(Trim$(record.DispatchRegistryNumber)) > 0 Then Exit Function
    If Len(Trim$(record.DispatchRegistryDate)) > 0 Then Exit Function

    IsDispatchRecordAvailable = True

End Function



Private Sub LoadSendersList()

    Set senderItems = DispatchRepositoryLoadSenders()

    cmbDispatchSender.Clear



    Dim i As Long

    For i = 1 To senderItems.count

        cmbDispatchSender.AddItem CStr(senderItems(i)(SenderColumnName))

    Next i

End Sub



Private Sub LoadEnvelopeFormatList()

    Set envelopeFormats = DispatchRepositoryLoadEnvelopeFormats()

    cmbDispatchEnvelopeFormat.Clear



    Dim i As Long

    For i = 1 To envelopeFormats.count

        cmbDispatchEnvelopeFormat.AddItem CStr(envelopeFormats(i)(EnvelopeFormatColumnDisplayName))

    Next i

End Sub

Private Sub LoadMailTypeList()

    DispatchRepositoryPopulateMailTypeOptions cmbDispatchMailType

End Sub



Private Sub SelectDefaultValues()

    If Not SelectEnvelopeFormatByKey("c5") Then
        If cmbDispatchEnvelopeFormat.ListCount > 0 Then cmbDispatchEnvelopeFormat.listIndex = 0
    End If



    Dim defaultSender As String

    defaultSender = DispatchRepositoryGetDefaultSenderName()

    If Len(defaultSender) > 0 Then

        SelectComboValue cmbDispatchSender, defaultSender

    ElseIf cmbDispatchSender.ListCount > 0 Then

        cmbDispatchSender.listIndex = 0

    End If



    If cmbDispatchMailType.ListCount > 0 Then cmbDispatchMailType.listIndex = 0

    ApplyWorkingRegistryDefaults



    If Len(Trim$(txtDispatchRegistryDate.Text)) = 0 Then

        txtDispatchRegistryDate.Text = Format$(Date, "dd.mm.yyyy")

    End If



    ApplyInitialSearchFocus

End Sub

Private Sub ApplyWorkingRegistryDefaults()
    Dim workingRegistryNumber As String
    workingRegistryNumber = DispatchRepositoryGetCurrentWorkingRegistryNumber()

    Dim workingRegistryDate As String
    workingRegistryDate = DispatchRepositoryGetCurrentWorkingRegistryDate()

    If Len(workingRegistryNumber) > 0 Or Len(workingRegistryDate) > 0 Then
        If Len(Trim$(txtDispatchRegistryNumber.Text)) = 0 And Len(Trim$(txtDispatchRegistryDate.Text)) = 0 Then
            If Not workingRegistryPromptShown Then
                workingRegistryPromptShown = True

                Dim promptText As String
                promptText = t("form.mail_dispatch.prompt.open_registry", "There is an open working registry. Click Yes to continue it. Click No to start a new registry.")

                Dim promptResult As VbMsgBoxResult
                promptResult = MsgBox(promptText, vbQuestion + vbYesNo, t("form.mail_dispatch.title", "Mail dispatch"))

                If promptResult = vbNo Then
                    txtDispatchRegistryNumber.Text = ""
                    txtDispatchRegistryDate.Text = Format$(Date, "dd.mm.yyyy")
                    Exit Sub
                End If
            End If
        End If
    End If

    If Len(Trim$(txtDispatchRegistryNumber.Text)) = 0 Then
        If Len(workingRegistryNumber) > 0 Then txtDispatchRegistryNumber.Text = workingRegistryNumber
    End If

    If Len(Trim$(txtDispatchRegistryDate.Text)) = 0 Then
        If Len(workingRegistryDate) > 0 Then txtDispatchRegistryDate.Text = workingRegistryDate
    End If
End Sub



Public Sub ApplyInitialSearchFocus()

    On Error Resume Next

    txtDispatchSearch.SetFocus

    On Error GoTo 0

End Sub



Private Sub lstDispatchLetters_Click()

    UpdateDispatchPreview

End Sub



Private Sub lstDispatchLetters_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Cancel = True

    QueueMailDispatchDoubleClick "available", lstDispatchLetters.listIndex

End Sub



Private Sub lstDispatchPackage_Click()

    UpdateDispatchPreview

End Sub



Private Sub btnDispatchRefresh_Click()

    Set packageLettersData = New Collection

    LoadDispatchData

    SelectDefaultValues

    UpdateDispatchPreview

End Sub



Private Sub btnDispatchClose_Click()

    Unload Me

End Sub



Private Sub btnDispatchCreate_Click()

    On Error GoTo CreateError



    If packageLettersData Is Nothing Or packageLettersData.count = 0 Then

        MsgBox t("form.mail_dispatch.error.no_package_items", "Put at least one letter into the envelope package."), vbExclamation

        Exit Sub

    End If



    If cmbDispatchSender.listIndex < 0 Then

        MsgBox t("form.mail_dispatch.error.no_sender", "Select a sender."), vbExclamation

        Exit Sub

    End If



    Dim envelopeFormatKey As String

    envelopeFormatKey = GetSelectedEnvelopeFormatKey()

    If Len(envelopeFormatKey) = 0 Then

        MsgBox t("form.mail_dispatch.error.no_envelope_format", "Select an envelope format."), vbExclamation

        Exit Sub

    End If



    If Len(Trim$(txtDispatchRegistryNumber.Text)) = 0 Then

        MsgBox t("form.mail_dispatch.error.no_registry_number", "Enter a registry number for the package."), vbExclamation

        Exit Sub

    End If



    If Not IsDateStringValid(txtDispatchRegistryDate.Text) Then

        MsgBox t("form.mail_dispatch.error.invalid_registry_date", "Enter a valid registry date in dd.mm.yyyy format."), vbExclamation

        Exit Sub

    End If



    Dim selectedSenderName As String
    selectedSenderName = cmbDispatchSender.Text

    Dim selectedRegistryNumber As String
    selectedRegistryNumber = txtDispatchRegistryNumber.Text

    Dim selectedRegistryDate As String
    selectedRegistryDate = txtDispatchRegistryDate.Text

    Dim selectedMailType As String
    selectedMailType = GetSelectedMailTypeKey()

    Dim selectedComment As String
    selectedComment = txtDispatchComment.Text

    Dim batchId As String
    batchId = DispatchRepositoryCreatePackageFromHistoryRecords(packageLettersData, selectedSenderName, envelopeFormatKey, selectedRegistryNumber, selectedRegistryDate, selectedMailType, "", "", selectedComment)



    If Len(batchId) = 0 Then

        MsgBox t("form.mail_dispatch.error.create_failed", "Failed to add the dispatch package to the worksheet."), vbCritical

        Exit Sub

    End If



    Dim registryRows As Long
    registryRows = BuildDispatchRegistry()

    Dim preparedEnvelopes As Long
    If registryRows > 0 Then preparedEnvelopes = PrepareEnvelopePrintForBatch(batchId)

    Dim resultMessage As String
    resultMessage = t("form.mail_dispatch.msg.package_created", "Пакет отправлений сохранен в рабочую таблицу.") & vbCrLf
    resultMessage = resultMessage & batchId & vbCrLf
    resultMessage = resultMessage & t("form.mail_dispatch.msg.letters_in_package", "Писем в пакете: ") & packageLettersData.count & vbCrLf
    resultMessage = resultMessage & t("form.mail_dispatch.msg.registry_rows", "Rows in current registry: ") & registryRows & vbCrLf
    resultMessage = resultMessage & t("form.mail_dispatch.msg.envelopes_prepared", "Prepared envelopes: ") & preparedEnvelopes

    MsgBox resultMessage, vbInformation

    If preparedEnvelopes > 0 Then
        If ShouldOpenPreparedEnvelopePreview() Then
            OpenPreparedEnvelopePreviewAndContinue batchId
            Exit Sub
        End If
    End If

    btnDispatchRefresh_Click

    Exit Sub



CreateError:

    MsgBox t("form.mail_dispatch.error.runtime", "Mail dispatch preparation error: ") & Err.description, vbCritical

End Sub

Private Sub FinalizeCurrentRegistry()

    On Error GoTo FinalizeError

    If Not ConfirmPostalRegistryPdfWithUnpackedLetters() Then Exit Sub

    BuildDispatchRegistry
    BuildPostalRegistryPrintSheet

    Dim pdfPath As String
    pdfPath = ExportPostalRegistryPrint()

    If Len(Trim$(pdfPath)) > 0 Then
        MsgBox t("form.mail_dispatch.msg.registry_finalized", "OPS registry exported and the mail batch was finalized.") & vbCrLf & pdfPath, vbInformation, t("postal.registry.pdf.title", "OPS registry")
        Unload Me
    Else
        MsgBox t("form.mail_dispatch.error.registry_finalize_failed", "OPS registry was not exported."), vbExclamation, t("postal.registry.pdf.title", "OPS registry")
    End If

    Exit Sub

FinalizeError:

    MsgBox t("form.mail_dispatch.error.registry_finalize_runtime", "OPS registry finalization error: ") & Err.description, vbCritical, t("postal.registry.pdf.title", "OPS registry")

End Sub

Private Function ShouldOpenPreparedEnvelopePreview() As Boolean
    ShouldOpenPreparedEnvelopePreview = MsgBox(t("form.mail_dispatch.prompt.preview_envelope", "Open envelope print preview now?"), vbQuestion + vbYesNo, t("form.mail_dispatch.title", "Mail dispatch")) = vbYes
End Function



Private Sub OpenPreparedEnvelopePreviewAndContinue(batchId As String)

    On Error GoTo PreviewError

    Me.Hide
    DoEvents
    PreviewPreparedEnvelopeForBatch batchId
    btnDispatchRefresh_Click
    Me.Show
    Exit Sub

PreviewError:

    MsgBox t("form.mail_dispatch.error.preview_failed", "Failed to open envelope print preview: ") & Err.description, vbExclamation
    btnDispatchRefresh_Click
    Me.Show

End Sub



Private Sub TransferSelectedLettersToPackage()

    On Error GoTo TransferError



    If lstDispatchLetters.ListCount = 0 Then Exit Sub



    Dim selectedIndexes As Collection

    Set selectedIndexes = GetSelectedIndexes(lstDispatchLetters)

    If selectedIndexes.count = 0 Then

        MsgBox t("form.mail_dispatch.error.no_letter", "Выберите письмо для подготовки отправления."), vbExclamation

        Exit Sub

    End If



    Dim targetPackageKey As String

    If packageLettersData.count > 0 Then

        targetPackageKey = BuildPackageGroupingKey(packageLettersData(1))

    End If



    Dim i As Long

    For i = 1 To selectedIndexes.count

        Dim listIndex As Long

        listIndex = CLng(selectedIndexes(i))



        Dim record As clsLetterHistoryRecord

        Set record = availableLettersData(listIndex + 1)



        If Len(targetPackageKey) > 0 Then

            If StrComp(BuildPackageGroupingKey(record), targetPackageKey, vbTextCompare) <> 0 Then

                MsgBox t("form.mail_dispatch.error.mixed_rpbs", "Letters in one envelope must have the same RPBS. If RPBS is empty, the addressee must match."), vbExclamation

                Exit Sub

            End If

        Else

            targetPackageKey = BuildPackageGroupingKey(record)

        End If

    Next i



    Dim recordsToMove As Collection

    Set recordsToMove = New Collection



    For i = 1 To selectedIndexes.count

        recordsToMove.Add availableLettersData(CLng(selectedIndexes(i)) + 1)

    Next i



    For i = selectedIndexes.count To 1 Step -1

        Dim removeIndex As Long

        removeIndex = CLng(selectedIndexes(i))

        RemoveHistoryRecordByKey allAvailableLettersData, BuildHistoryRecordKey(recordsToMove(i))

    Next i



    For i = 1 To recordsToMove.count

        packageLettersData.Add recordsToMove(i)

    Next i



    ApplyAvailableLettersFilter

    RebindPackageLettersList

    UpdateDispatchPreview

    Exit Sub



TransferError:

    MsgBox t("form.mail_dispatch.error.transfer_failed", "Не удалось переместить выбранные письма в пакет."), vbExclamation

End Sub



Private Sub TransferCurrentLetterToPackage()

    On Error GoTo TransferError

    If lstDispatchLetters.ListCount = 0 Then Exit Sub

    If lstDispatchLetters.listIndex < 0 Then Exit Sub

    If lstDispatchLetters.listIndex + 1 > availableLettersData.count Then Exit Sub

    Dim listIndex As Long

    listIndex = lstDispatchLetters.listIndex

    Dim record As clsLetterHistoryRecord

    Set record = availableLettersData(listIndex + 1)

    If Not CanAddRecordToCurrentPackage(record) Then Exit Sub

    RemoveHistoryRecordByKey allAvailableLettersData, BuildHistoryRecordKey(record)

    availableLettersData.Remove listIndex + 1

    packageLettersData.Add record

    lstDispatchLetters.RemoveItem listIndex

    ClearListSelection lstDispatchLetters

    RebindPackageLettersList

    UpdateDispatchPreview

    Exit Sub

TransferError:

    MsgBox t("form.mail_dispatch.error.transfer_failed", "Не удалось переместить выбранные письма в пакет."), vbExclamation

End Sub



Private Function CanAddRecordToCurrentPackage(record As clsLetterHistoryRecord) As Boolean

    CanAddRecordToCurrentPackage = False

    Dim targetPackageKey As String

    If packageLettersData.count > 0 Then

        targetPackageKey = BuildPackageGroupingKey(packageLettersData(1))

    End If

    If Len(targetPackageKey) > 0 Then

        If StrComp(BuildPackageGroupingKey(record), targetPackageKey, vbTextCompare) <> 0 Then

            MsgBox t("form.mail_dispatch.error.mixed_rpbs", "Letters in one envelope must have the same RPBS. If RPBS is empty, the addressee must match."), vbExclamation

            Exit Function

        End If

    End If

    CanAddRecordToCurrentPackage = True

End Function

Private Function BuildPackageGroupingKey(record As clsLetterHistoryRecord) As String

    If record Is Nothing Then Exit Function

    BuildPackageGroupingKey = DispatchRepositoryResolvePackageGroupingKey(record.Addressee)

End Function



Private Sub RemoveSelectedLettersFromPackage()

    On Error GoTo TransferError



    If lstDispatchPackage.ListCount = 0 Then Exit Sub



    Dim selectedIndexes As Collection

    Set selectedIndexes = GetSelectedIndexes(lstDispatchPackage)

    If selectedIndexes.count = 0 Then Exit Sub



    Dim i As Long

    Dim recordsToReturn As Collection

    Set recordsToReturn = New Collection



    For i = 1 To selectedIndexes.count

        recordsToReturn.Add packageLettersData(CLng(selectedIndexes(i)) + 1)

    Next i



    For i = selectedIndexes.count To 1 Step -1

        Dim removeIndex As Long

        removeIndex = CLng(selectedIndexes(i))

        packageLettersData.Remove removeIndex + 1

    Next i



    For i = 1 To recordsToReturn.count

        AddHistoryRecordIfMissing allAvailableLettersData, recordsToReturn(i)

    Next i



    ApplyAvailableLettersFilter

    RebindPackageLettersList

    UpdateDispatchPreview

    Exit Sub



TransferError:

    MsgBox t("form.mail_dispatch.error.transfer_back_failed", "Не удалось вернуть письма из пакета."), vbExclamation

End Sub



Private Sub RemoveCurrentLetterFromPackage()

    On Error GoTo TransferError

    If lstDispatchPackage.ListCount = 0 Then Exit Sub

    If lstDispatchPackage.listIndex < 0 Then Exit Sub

    If lstDispatchPackage.listIndex + 1 > packageLettersData.count Then Exit Sub

    Dim listIndex As Long

    listIndex = lstDispatchPackage.listIndex

    Dim record As clsLetterHistoryRecord

    Set record = packageLettersData(listIndex + 1)

    packageLettersData.Remove listIndex + 1

    AddHistoryRecordIfMissing allAvailableLettersData, record

    lstDispatchPackage.RemoveItem listIndex

    ClearListSelection lstDispatchPackage

    ApplyAvailableLettersFilter

    UpdateDispatchPreview

    Exit Sub

TransferError:

    MsgBox t("form.mail_dispatch.error.transfer_failed", "Не удалось вернуть письмо из пакета."), vbExclamation

End Sub



Private Sub UpdateDispatchPreview()

    If packageLettersData.count > 0 Then

        txtDispatchPreview.Text = BuildPackagePreviewText()

        Exit Sub

    End If



    Dim record As clsLetterHistoryRecord

    Set record = GetSelectedAvailableHistoryRecord()



    If record Is Nothing Then

        txtDispatchPreview.Text = ""

        Exit Sub

    End If



    txtDispatchPreview.Text = DispatchRepositoryBuildRecipientPreviewByAddressee(record.Addressee)

End Sub



Private Function GetSelectedAvailableHistoryRecord() As clsLetterHistoryRecord

    If lstDispatchLetters.listIndex < 0 Then Exit Function

    If availableLettersData Is Nothing Then Exit Function

    If lstDispatchLetters.listIndex + 1 > availableLettersData.count Then Exit Function



    If TypeName(availableLettersData(lstDispatchLetters.listIndex + 1)) = "clsLetterHistoryRecord" Then

        Set GetSelectedAvailableHistoryRecord = availableLettersData(lstDispatchLetters.listIndex + 1)

    End If

End Function



Private Function GetSelectedEnvelopeFormatKey() As String

    If cmbDispatchEnvelopeFormat.listIndex < 0 Then Exit Function

    If envelopeFormats Is Nothing Then Exit Function

    If cmbDispatchEnvelopeFormat.listIndex + 1 > envelopeFormats.count Then Exit Function



    GetSelectedEnvelopeFormatKey = CStr(envelopeFormats(cmbDispatchEnvelopeFormat.listIndex + 1)(EnvelopeFormatColumnKey))

End Function

Private Function SelectEnvelopeFormatByKey(envelopeFormatKey As String) As Boolean

    If envelopeFormats Is Nothing Then Exit Function

    Dim normalizedKey As String
    normalizedKey = LCase$(Trim$(envelopeFormatKey))

    Dim i As Long
    For i = 1 To envelopeFormats.count
        If LCase$(Trim$(CStr(envelopeFormats(i)(EnvelopeFormatColumnKey)))) = normalizedKey Then
            cmbDispatchEnvelopeFormat.listIndex = i - 1
            SelectEnvelopeFormatByKey = True
            Exit Function
        End If
    Next i

End Function

Private Function GetSelectedMailTypeKey() As String

    GetSelectedMailTypeKey = DispatchRepositoryNormalizeMailTypeKey(cmbDispatchMailType.Text)

End Function



Private Sub SelectComboValue(targetCombo As ComboBox, expectedValue As String)

    Dim i As Long

    For i = 0 To targetCombo.ListCount - 1

        If StrComp(CStr(targetCombo.List(i)), expectedValue, vbTextCompare) = 0 Then

            targetCombo.listIndex = i

            Exit Sub

        End If

    Next i

End Sub



Private Sub SetLocalizedCaption(controlName As String, translationKey As String, fallbackText As String)

    On Error Resume Next

    Me.Controls(controlName).Caption = t(translationKey, fallbackText)

    On Error GoTo 0

End Sub



Private Sub EnsureDynamicControls()

    Set lblDispatchSearch = EnsureDynamicLabel("lblDispatchSearch")

    Set txtDispatchSearch = EnsureDynamicTextBox("txtDispatchSearch")

    Set lblDispatchPackage = EnsureDynamicLabel("lblDispatchPackage")

    Set lstDispatchPackage = EnsureDynamicListBox("lstDispatchPackage")

    Set btnDispatchAddToPackage = EnsureDynamicButton("btnDispatchAddToPackage")

    Set btnDispatchRemoveFromPackage = EnsureDynamicButton("btnDispatchRemoveFromPackage")

    Set btnDispatchFinalizeRegistry = EnsureDynamicButton("btnDispatchFinalizeRegistry")

    Set lblDispatchRegistryNumber = EnsureDynamicLabel("lblDispatchRegistryNumber")

    Set txtDispatchRegistryNumber = EnsureDynamicTextBox("txtDispatchRegistryNumber")

    Set lblDispatchRegistryDate = EnsureDynamicLabel("lblDispatchRegistryDate")

    Set txtDispatchRegistryDate = EnsureDynamicTextBox("txtDispatchRegistryDate")

    Set cmbDispatchMailType = EnsureDynamicComboBox("cmbDispatchMailType")

    BindDynamicButtonHandlers

End Sub



Private Function EnsureDynamicLabel(controlName As String) As MSForms.Label

    If ControlExists(controlName) Then

        Set EnsureDynamicLabel = Me.Controls(controlName)

    Else

        Set EnsureDynamicLabel = Me.Controls.Add("Forms.Label.1", controlName, True)

    End If

End Function

Private Function EnsureDynamicComboBox(controlName As String) As MSForms.ComboBox

    If ControlExists(controlName) Then

        Set EnsureDynamicComboBox = Me.Controls(controlName)

    Else

        Set EnsureDynamicComboBox = Me.Controls.Add("Forms.ComboBox.1", controlName, True)

    End If

End Function



Private Function EnsureDynamicTextBox(controlName As String) As MSForms.TextBox

    If ControlExists(controlName) Then

        Set EnsureDynamicTextBox = Me.Controls(controlName)

    Else

        Set EnsureDynamicTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)

    End If

End Function



Private Function EnsureDynamicListBox(controlName As String) As MSForms.ListBox

    If ControlExists(controlName) Then

        Set EnsureDynamicListBox = Me.Controls(controlName)

    Else

        Set EnsureDynamicListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)

    End If

End Function



Private Function EnsureDynamicButton(controlName As String) As MSForms.CommandButton

    If ControlExists(controlName) Then

        Set EnsureDynamicButton = Me.Controls(controlName)

    Else

        Set EnsureDynamicButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)

    End If

End Function



Private Function ControlExists(controlName As String) As Boolean

    On Error Resume Next

    ControlExists = Not Me.Controls(controlName) Is Nothing

    On Error GoTo 0

End Function



Private Sub BindDynamicButtonHandlers()

    Set dynamicButtonHandlers = New Collection



    Dim addHandler As clsDispatchDynamicButtonHandler

    Set addHandler = New clsDispatchDynamicButtonHandler

    addHandler.BindButton btnDispatchAddToPackage, Me

    dynamicButtonHandlers.Add addHandler



    Dim removeHandler As clsDispatchDynamicButtonHandler

    Set removeHandler = New clsDispatchDynamicButtonHandler

    removeHandler.BindButton btnDispatchRemoveFromPackage, Me

    dynamicButtonHandlers.Add removeHandler



    Dim finalizeHandler As clsDispatchDynamicButtonHandler

    Set finalizeHandler = New clsDispatchDynamicButtonHandler

    finalizeHandler.BindButton btnDispatchFinalizeRegistry, Me

    dynamicButtonHandlers.Add finalizeHandler



    Dim searchHandler As clsDispatchDynamicButtonHandler

    Set searchHandler = New clsDispatchDynamicButtonHandler

    searchHandler.BindTextBox txtDispatchSearch, Me

    dynamicButtonHandlers.Add searchHandler



    Dim packageListHandler As clsDispatchDynamicButtonHandler

    Set packageListHandler = New clsDispatchDynamicButtonHandler

    packageListHandler.BindListBox lstDispatchPackage, Me

    dynamicButtonHandlers.Add packageListHandler

End Sub



Private Sub RebindAvailableLettersList()

    lstDispatchLetters.Clear



    Dim i As Long

    For i = 1 To availableLettersData.count

        lstDispatchLetters.AddItem RepositoryFormatLetterHistoryDisplay(availableLettersData(i))

    Next i



    If lstDispatchLetters.ListCount > 0 Then

        lstDispatchLetters.listIndex = 0

    End If


    ClearListSelection lstDispatchLetters
End Sub



Private Sub ApplyAvailableLettersFilter()

    Set availableLettersData = RepositoryFilterLetterHistoryRecords(allAvailableLettersData, txtDispatchSearch.Text)

    RebindAvailableLettersList

End Sub



Private Sub RebindPackageLettersList()

    lstDispatchPackage.Clear



    Dim i As Long

    For i = 1 To packageLettersData.count

        lstDispatchPackage.AddItem RepositoryFormatLetterHistoryDisplay(packageLettersData(i))

    Next i


    ClearListSelection lstDispatchPackage
End Sub



Private Sub ClearListSelection(targetList As MSForms.ListBox)

    Dim i As Long

    For i = 0 To targetList.ListCount - 1

        targetList.Selected(i) = False

    Next i

    targetList.listIndex = -1

End Sub


Private Sub SelectSingleListIndex(targetList As MSForms.ListBox, ByVal listIndex As Long)

    ClearListSelection targetList

    If listIndex < 0 Then Exit Sub

    If listIndex >= targetList.ListCount Then Exit Sub

    targetList.Selected(listIndex) = True

    targetList.listIndex = listIndex

End Sub



Private Function GetSelectedIndexes(targetList As MSForms.ListBox) As Collection

    Set GetSelectedIndexes = New Collection



    Dim i As Long

    For i = 0 To targetList.ListCount - 1

        If targetList.Selected(i) Then

            GetSelectedIndexes.Add i

        End If

    Next i

End Function



Private Function BuildHistoryRecordKey(record As clsLetterHistoryRecord) As String

    BuildHistoryRecordKey = UCase$(Trim$(record.Addressee)) & "|" & UCase$(Trim$(record.OutgoingNumber)) & "|" & UCase$(Trim$(record.OutgoingDate))

End Function



Private Sub RemoveHistoryRecordByKey(targetRecords As Collection, recordKey As String)

    Dim i As Long

    For i = targetRecords.count To 1 Step -1

        If BuildHistoryRecordKey(targetRecords(i)) = recordKey Then

            targetRecords.Remove i

            Exit Sub

        End If

    Next i

End Sub



Private Sub AddHistoryRecordIfMissing(targetRecords As Collection, record As clsLetterHistoryRecord)

    Dim recordKey As String

    recordKey = BuildHistoryRecordKey(record)



    Dim i As Long

    For i = 1 To targetRecords.count

        If BuildHistoryRecordKey(targetRecords(i)) = recordKey Then

            Exit Sub

        End If

    Next i



    targetRecords.Add record

End Sub



Private Function IsDateStringValid(dateText As String) As Boolean

    On Error GoTo InvalidDate

    If Len(Trim$(dateText)) = 0 Then Exit Function

    IsDateStringValid = IsDate(CDate(dateText))

    Exit Function

InvalidDate:

    IsDateStringValid = False

End Function



Private Function BuildPackagePreviewText() As String

    Dim firstRecord As clsLetterHistoryRecord

    Set firstRecord = packageLettersData(1)



    BuildPackagePreviewText = DispatchRepositoryBuildRecipientPreviewByAddressee(firstRecord.Addressee)

    Dim rpbsCode As String

    rpbsCode = DispatchRepositoryGetRpbsByAddressee(firstRecord.Addressee)

    If Len(rpbsCode) > 0 Then

        BuildPackagePreviewText = BuildPackagePreviewText & vbCrLf & t("form.letter_creator.label.rpbs", "RPBS") & ": " & rpbsCode

    End If

    BuildPackagePreviewText = BuildPackagePreviewText & vbCrLf & vbCrLf & BuildOutgoingNumbersText(packageLettersData)



    If Len(Trim$(cmbDispatchSender.Text)) > 0 Then

        BuildPackagePreviewText = BuildPackagePreviewText & vbCrLf & vbCrLf & t("form.mail_dispatch.preview.sender", "Отправитель:") & " " & cmbDispatchSender.Text

    End If



    If Len(Trim$(txtDispatchRegistryNumber.Text)) > 0 Then

        BuildPackagePreviewText = BuildPackagePreviewText & vbCrLf & t("form.mail_dispatch.preview.registry", "Реестр:") & " " & txtDispatchRegistryNumber.Text

    End If

End Function



Private Function BuildOutgoingNumbersText(records As Collection) As String

    Dim i As Long

    For i = 1 To records.count

        If i > 1 Then

            BuildOutgoingNumbersText = BuildOutgoingNumbersText & vbCrLf

        End If

        BuildOutgoingNumbersText = BuildOutgoingNumbersText & BuildOutgoingLine(records(i))

    Next i

End Function



Private Function BuildOutgoingLine(record As clsLetterHistoryRecord) As String

    BuildOutgoingLine = Trim$(record.OutgoingNumber)



    If Len(Trim$(record.OutgoingDate)) > 0 Then

        BuildOutgoingLine = BuildOutgoingLine & " " & t("common.preposition.from", "от") & " " & Trim$(record.OutgoingDate)

    End If

End Function



Public Sub HandleDynamicButtonClick(controlName As String)

    Select Case controlName

    Case "btnDispatchAddToPackage"

        TransferSelectedLettersToPackage

    Case "btnDispatchRemoveFromPackage"

        RemoveSelectedLettersFromPackage

    Case "btnDispatchFinalizeRegistry"

        FinalizeCurrentRegistry

    End Select

End Sub



Public Sub HandleDynamicTextChanged(controlName As String)

    Select Case controlName

    Case "txtDispatchSearch"

        ApplyAvailableLettersFilter

        UpdateDispatchPreview

    End Select

End Sub



Public Sub HandleDynamicListDoubleClick(controlName As String)

    Select Case controlName

    Case "lstDispatchPackage"

        QueueMailDispatchDoubleClick "package", lstDispatchPackage.listIndex

    End Select

End Sub



Private Sub QueueMailDispatchDoubleClick(actionName As String, ByVal listIndex As Long)

    On Error GoTo QueueError

    If listIndex < 0 Then Exit Sub

    pendingDoubleClickAction = actionName
    pendingDoubleClickIndex = listIndex

    If doubleClickActionScheduled Then Exit Sub

    doubleClickActionScheduled = True
    pendingDoubleClickRunAt = Now + TimeValue("00:00:01")
    Application.OnTime pendingDoubleClickRunAt, "RunMailDispatchDeferredDoubleClick"
    Exit Sub

QueueError:

    ClearPendingDoubleClickState
    MsgBox t("form.mail_dispatch.error.transfer_failed", "Не удалось переместить выбранные письма в пакет."), vbExclamation

End Sub



Public Sub RunDeferredDoubleClickAction()

    On Error GoTo DeferredError

    doubleClickActionScheduled = False
    pendingDoubleClickRunAt = 0

    Select Case pendingDoubleClickAction

    Case "available"

        SelectSingleListIndex lstDispatchLetters, pendingDoubleClickIndex
        TransferSelectedLettersToPackage

    Case "package"

        SelectSingleListIndex lstDispatchPackage, pendingDoubleClickIndex
        RemoveSelectedLettersFromPackage

    End Select

    ClearPendingDoubleClickState
    Exit Sub

DeferredError:

    ClearPendingDoubleClickState
    MsgBox t("form.mail_dispatch.error.transfer_failed", "Не удалось переместить выбранные письма в пакет."), vbExclamation

End Sub



Private Sub CancelPendingDoubleClickSchedule()

    On Error Resume Next

    If doubleClickActionScheduled Then
        If pendingDoubleClickRunAt <> 0 Then
            Application.OnTime EarliestTime:=pendingDoubleClickRunAt, Procedure:="RunMailDispatchDeferredDoubleClick", Schedule:=False
        End If
    End If

    On Error GoTo 0
    ClearPendingDoubleClickState

End Sub



Private Sub ClearPendingDoubleClickState()

    doubleClickActionScheduled = False
    pendingDoubleClickAction = ""
    pendingDoubleClickIndex = -1
    pendingDoubleClickRunAt = 0

End Sub


