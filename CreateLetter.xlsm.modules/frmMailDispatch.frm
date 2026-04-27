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

' Form: frmMailDispatch v1.2.1

' Author: CreateLetter contributors

' Date: 27.04.2026

' Purpose: Thin-shell UI for preparing grouped dispatch packages from existing letters

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

Private lblDispatchRegistryNumber As MSForms.Label

Private txtDispatchRegistryNumber As MSForms.TextBox

Private lblDispatchRegistryDate As MSForms.Label

Private txtDispatchRegistryDate As MSForms.TextBox



Private Sub UserForm_Initialize()

    Set allAvailableLettersData = New Collection

    Set availableLettersData = New Collection

    Set packageLettersData = New Collection

    Set envelopeFormats = New Collection

    Set senderItems = New Collection

    Set dynamicButtonHandlers = New Collection



    EnsureDynamicControls

    ApplyFormSettings

    ApplyResponsiveLayout

    ApplyLocalizedCaptions

    ConfigureLists

    LoadDispatchData

    SelectDefaultValues

    UpdateDispatchPreview

End Sub



Private Sub ApplyFormSettings()

    With Me

        .Caption = t("form.mail_dispatch.title", "Mail dispatch") & " v1.2.0"

        .backColor = RGB(248, 248, 248)

    End With

End Sub



Private Sub ApplyResponsiveLayout()

    Const FORM_WIDTH As Single = 920

    Const FORM_HEIGHT As Single = 700

    Const LEFT_COLUMN_LEFT As Single = 18

    Const PACKAGE_COLUMN_LEFT As Single = 472

    Const LIST_WIDTH As Single = 330

    Const CONTENT_TOP As Single = 36

    Const SEARCH_TOP As Single = 66

    Const LIST_TOP As Single = 116

    Const LIST_HEIGHT As Single = 264

    Const MIDDLE_BUTTON_LEFT As Single = 380

    Const BUTTON_TOP As Single = 180

    Const METADATA_TOP As Single = 402

    Const COMMENT_TOP As Single = 522

    Const PREVIEW_TOP As Single = 522

    Const PREVIEW_HEIGHT As Single = 126



    Me.StartUpPosition = 1

    Me.Width = FORM_WIDTH

    Me.Height = FORM_HEIGHT

    Me.ScrollBars = fmScrollBarsNone



    lblDispatchLetters.Left = LEFT_COLUMN_LEFT

    lblDispatchLetters.Top = CONTENT_TOP

    lblDispatchLetters.Width = LIST_WIDTH



    lblDispatchSearch.Left = LEFT_COLUMN_LEFT

    lblDispatchSearch.Top = SEARCH_TOP

    lblDispatchSearch.Width = 200



    txtDispatchSearch.Left = LEFT_COLUMN_LEFT

    txtDispatchSearch.Top = SEARCH_TOP + 20

    txtDispatchSearch.Width = LIST_WIDTH

    txtDispatchSearch.Height = 22



    lstDispatchLetters.Left = LEFT_COLUMN_LEFT

    lstDispatchLetters.Top = LIST_TOP

    lstDispatchLetters.Width = LIST_WIDTH

    lstDispatchLetters.Height = LIST_HEIGHT



    btnDispatchRefresh.Left = LEFT_COLUMN_LEFT

    btnDispatchRefresh.Top = COMMENT_TOP + 116

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

    btnDispatchAddToPackage.Width = 44

    btnDispatchAddToPackage.Height = 28



    btnDispatchRemoveFromPackage.Left = MIDDLE_BUTTON_LEFT

    btnDispatchRemoveFromPackage.Top = BUTTON_TOP + 40

    btnDispatchRemoveFromPackage.Width = 44

    btnDispatchRemoveFromPackage.Height = 28



    lblDispatchSender.Left = LEFT_COLUMN_LEFT

    lblDispatchSender.Top = METADATA_TOP

    cmbDispatchSender.Left = LEFT_COLUMN_LEFT

    cmbDispatchSender.Top = METADATA_TOP + 22

    cmbDispatchSender.Width = 220



    lblDispatchEnvelopeFormat.Left = 260

    lblDispatchEnvelopeFormat.Top = METADATA_TOP

    cmbDispatchEnvelopeFormat.Left = 260

    cmbDispatchEnvelopeFormat.Top = METADATA_TOP + 22

    cmbDispatchEnvelopeFormat.Width = 86



    lblDispatchMailType.Left = 370

    lblDispatchMailType.Top = METADATA_TOP

    txtDispatchMailType.Left = 370

    txtDispatchMailType.Top = METADATA_TOP + 22

    txtDispatchMailType.Width = 140



    lblDispatchRegistryNumber.Left = LEFT_COLUMN_LEFT

    lblDispatchRegistryNumber.Top = METADATA_TOP + 62

    txtDispatchRegistryNumber.Left = LEFT_COLUMN_LEFT

    txtDispatchRegistryNumber.Top = METADATA_TOP + 84

    txtDispatchRegistryNumber.Width = 120



    lblDispatchRegistryDate.Left = 160

    lblDispatchRegistryDate.Top = METADATA_TOP + 62

    txtDispatchRegistryDate.Left = 160

    txtDispatchRegistryDate.Top = METADATA_TOP + 84

    txtDispatchRegistryDate.Width = 120



    lblDispatchMass.Visible = False

    txtDispatchMass.Visible = False

    lblDispatchDeclaredValue.Visible = False

    txtDispatchDeclaredValue.Visible = False



    lblDispatchComment.Left = LEFT_COLUMN_LEFT

    lblDispatchComment.Top = COMMENT_TOP - 22

    txtDispatchComment.Left = LEFT_COLUMN_LEFT

    txtDispatchComment.Top = COMMENT_TOP

    txtDispatchComment.Width = 360

    txtDispatchComment.Height = 48



    lblDispatchPreview.Left = 400

    lblDispatchPreview.Top = PREVIEW_TOP - 22

    txtDispatchPreview.Left = 400

    txtDispatchPreview.Top = PREVIEW_TOP

    txtDispatchPreview.Width = 360

    txtDispatchPreview.Height = PREVIEW_HEIGHT



    btnDispatchCreate.Left = 400

    btnDispatchCreate.Top = 620

    btnDispatchCreate.Width = 160



    btnDispatchClose.Left = 580

    btnDispatchClose.Top = 620

    btnDispatchClose.Width = 140

End Sub



Private Sub ApplyLocalizedCaptions()

    SetLocalizedCaption "lblDispatchLetters", "form.mail_dispatch.label.available_letters", "Доступные письма"

    lblDispatchSearch.Caption = t("form.mail_dispatch.label.search_letters", "Поиск писем")

    SetLocalizedCaption "lblDispatchSender", "form.mail_dispatch.label.sender", "Отправитель"

    SetLocalizedCaption "lblDispatchEnvelopeFormat", "form.mail_dispatch.label.envelope_format", "Формат конверта"

    SetLocalizedCaption "lblDispatchMailType", "form.mail_dispatch.label.mail_type", "Вид отправления"

    SetLocalizedCaption "lblDispatchRegistryNumber", "form.mail_dispatch.label.registry_number", "Номер реестра"

    SetLocalizedCaption "lblDispatchRegistryDate", "form.mail_dispatch.label.registry_date", "Дата реестра"

    SetLocalizedCaption "lblDispatchComment", "form.mail_dispatch.label.comment", "Комментарий"

    SetLocalizedCaption "lblDispatchPreview", "form.mail_dispatch.label.preview", "Предпросмотр"

    lblDispatchPackage.Caption = t("form.mail_dispatch.label.package_letters", "Пакет отправки")



    btnDispatchRefresh.Caption = t("form.mail_dispatch.button.refresh", "Обновить")

    btnDispatchCreate.Caption = t("form.mail_dispatch.button.create_package", "Сохранить пакет")

    btnDispatchClose.Caption = t("form.mail_dispatch.button.close", "Закрыть")

    btnDispatchAddToPackage.Caption = ">>"

    btnDispatchRemoveFromPackage.Caption = "<<"



    txtDispatchMailType.ControlTipText = t("form.mail_dispatch.tip.mail_type", "Например: заказное, простое, с уведомлением")

    txtDispatchSearch.ControlTipText = t("form.mail_dispatch.tip.search_letters", "Введите номер, дату, адресата или текст письма для фильтрации списка")

    txtDispatchComment.ControlTipText = t("form.mail_dispatch.tip.comment", "Короткий служебный комментарий для отправления")

    txtDispatchRegistryNumber.ControlTipText = t("form.mail_dispatch.tip.registry_number", "Номер внутреннего реестра для этого пакета")

    txtDispatchRegistryDate.ControlTipText = t("form.mail_dispatch.tip.registry_date", "Дата внутреннего реестра в формате дд.мм.гггг")

End Sub



Private Sub ConfigureLists()

    lstDispatchLetters.MultiSelect = fmMultiSelectMulti

    lstDispatchPackage.MultiSelect = fmMultiSelectMulti

    lstDispatchLetters.Font.Size = 10

    lstDispatchPackage.Font.Size = 10

    txtDispatchPreview.MultiLine = True

    txtDispatchPreview.ScrollBars = fmScrollBarsVertical

End Sub



Private Sub LoadDispatchData()

    LoadLettersList

    LoadSendersList

    LoadEnvelopeFormatList

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



Private Sub SelectDefaultValues()

    If cmbDispatchEnvelopeFormat.ListCount > 0 Then

        cmbDispatchEnvelopeFormat.listIndex = 0

    End If



    Dim defaultSender As String

    defaultSender = DispatchRepositoryGetDefaultSenderName()

    If Len(defaultSender) > 0 Then

        SelectComboValue cmbDispatchSender, defaultSender

    ElseIf cmbDispatchSender.ListCount > 0 Then

        cmbDispatchSender.listIndex = 0

    End If



    If Len(Trim$(txtDispatchMailType.Text)) = 0 Then

        txtDispatchMailType.Text = t("form.mail_dispatch.default.mail_type", "заказное")

    End If



    If Len(Trim$(txtDispatchRegistryDate.Text)) = 0 Then

        txtDispatchRegistryDate.Text = Format$(Date, "dd.mm.yyyy")

    End If



    ApplyInitialSearchFocus

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

    SelectSingleListIndex lstDispatchLetters, lstDispatchLetters.listIndex

    TransferCurrentLetterToPackage

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

        MsgBox t("form.mail_dispatch.error.no_package_items", "Добавьте хотя бы одно письмо в пакет отправки."), vbExclamation

        Exit Sub

    End If



    If cmbDispatchSender.listIndex < 0 Then

        MsgBox t("form.mail_dispatch.error.no_sender", "Выберите отправителя."), vbExclamation

        Exit Sub

    End If



    Dim envelopeFormatKey As String

    envelopeFormatKey = GetSelectedEnvelopeFormatKey()

    If Len(envelopeFormatKey) = 0 Then

        MsgBox t("form.mail_dispatch.error.no_envelope_format", "Выберите формат конверта."), vbExclamation

        Exit Sub

    End If



    If Len(Trim$(txtDispatchRegistryNumber.Text)) = 0 Then

        MsgBox t("form.mail_dispatch.error.no_registry_number", "Укажите номер реестра для пакета."), vbExclamation

        Exit Sub

    End If



    If Not IsDateStringValid(txtDispatchRegistryDate.Text) Then

        MsgBox t("form.mail_dispatch.error.invalid_registry_date", "Укажите корректную дату реестра в формате дд.мм.гггг."), vbExclamation

        Exit Sub

    End If



    Dim batchId As String

    batchId = DispatchRepositoryCreatePackageFromHistoryRecords( _

        packageLettersData, _

        cmbDispatchSender.Text, _

        envelopeFormatKey, _

        txtDispatchRegistryNumber.Text, _

        txtDispatchRegistryDate.Text, _

        txtDispatchMailType.Text, _

        "", _

        "", _

        txtDispatchComment.Text)



    If Len(batchId) = 0 Then

        MsgBox t("form.mail_dispatch.error.create_failed", "Не удалось добавить отправление в рабочую таблицу."), vbCritical

        Exit Sub

    End If



    MsgBox t("form.mail_dispatch.msg.package_created", "Пакет отправлений сохранен в рабочую таблицу.") & vbCrLf & _

           batchId & vbCrLf & _

           t("form.mail_dispatch.msg.letters_in_package", "Писем в пакете: ") & packageLettersData.count, vbInformation

    btnDispatchRefresh_Click

    Exit Sub



CreateError:

    MsgBox t("form.mail_dispatch.error.runtime", "Ошибка подготовки почтового отправления: ") & Err.description, vbCritical

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



    Dim targetAddressee As String

    If packageLettersData.count > 0 Then

        targetAddressee = packageLettersData(1).Addressee

    End If



    Dim i As Long

    For i = 1 To selectedIndexes.count

        Dim listIndex As Long

        listIndex = CLng(selectedIndexes(i))



        Dim record As clsLetterHistoryRecord

        Set record = availableLettersData(listIndex + 1)



        If Len(targetAddressee) > 0 Then

            If StrComp(Trim$(record.Addressee), Trim$(targetAddressee), vbTextCompare) <> 0 Then

                MsgBox t("form.mail_dispatch.error.mixed_addressee", "В один пакет можно добавлять только письма одному адресату."), vbExclamation

                Exit Sub

            End If

        Else

            targetAddressee = record.Addressee

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

    Dim targetAddressee As String

    If packageLettersData.count > 0 Then

        targetAddressee = packageLettersData(1).Addressee

    End If

    If Len(targetAddressee) > 0 Then

        If StrComp(Trim$(record.Addressee), Trim$(targetAddressee), vbTextCompare) <> 0 Then

            MsgBox t("form.mail_dispatch.error.mixed_addressee", "В один пакет можно добавлять только письма одному адресату."), vbExclamation

            Exit Function

        End If

    End If

    CanAddRecordToCurrentPackage = True

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

    Set lblDispatchRegistryNumber = EnsureDynamicLabel("lblDispatchRegistryNumber")

    Set txtDispatchRegistryNumber = EnsureDynamicTextBox("txtDispatchRegistryNumber")

    Set lblDispatchRegistryDate = EnsureDynamicLabel("lblDispatchRegistryDate")

    Set txtDispatchRegistryDate = EnsureDynamicTextBox("txtDispatchRegistryDate")

    BindDynamicButtonHandlers

End Sub



Private Function EnsureDynamicLabel(controlName As String) As MSForms.Label

    If ControlExists(controlName) Then

        Set EnsureDynamicLabel = Me.Controls(controlName)

    Else

        Set EnsureDynamicLabel = Me.Controls.Add("Forms.Label.1", controlName, True)

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

        SelectSingleListIndex lstDispatchPackage, lstDispatchPackage.listIndex

        RemoveCurrentLetterFromPackage

    End Select

End Sub


