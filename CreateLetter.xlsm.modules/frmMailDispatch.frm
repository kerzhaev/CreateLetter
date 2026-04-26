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
' Form: frmMailDispatch v1.0.0
' Author: CreateLetter contributors
' Date: 26.04.2026
' Purpose: Thin-shell UI for preparing dispatch items from existing letters
' ======================================================================
Option Explicit

Private allLettersData As Collection
Private envelopeFormats As Collection
Private senderItems As Collection

Private Sub UserForm_Initialize()
    Set allLettersData = New Collection
    Set envelopeFormats = New Collection
    Set senderItems = New Collection

    ApplyFormSettings
    ApplyResponsiveLayout
    ApplyLocalizedCaptions
    LoadDispatchData
    SelectDefaultValues
    UpdateDispatchPreview
End Sub

Private Sub ApplyFormSettings()
    With Me
        .Caption = t("form.mail_dispatch.title", "Mail dispatch") & " v1.0.0"
        .BackColor = RGB(248, 248, 248)
    End With
End Sub

Private Sub ApplyResponsiveLayout()
    Const FORM_WIDTH As Single = 648
    Const FORM_HEIGHT As Single = 468
    Const MIN_MARGIN As Single = 18

    Dim usableWidth As Single
    Dim usableHeight As Single
    usableWidth = Application.UsableWidth
    usableHeight = Application.UsableHeight

    Me.StartUpPosition = 0

    If usableWidth >= FORM_WIDTH + (MIN_MARGIN * 2) Then
        Me.Width = FORM_WIDTH
        Me.Left = Application.Left + ((usableWidth - Me.Width) / 2)
    Else
        Me.Width = usableWidth - (MIN_MARGIN * 2)
        Me.Left = Application.Left + MIN_MARGIN
    End If

    If usableHeight >= FORM_HEIGHT + (MIN_MARGIN * 2) Then
        Me.Height = FORM_HEIGHT
        Me.Top = Application.Top + ((usableHeight - Me.Height) / 2)
    Else
        Me.Height = usableHeight - (MIN_MARGIN * 2)
        Me.Top = Application.Top + MIN_MARGIN
        Me.ScrollBars = fmScrollBarsVertical
        Me.KeepScrollBarsVisible = fmScrollBarsVertical
        Me.ScrollHeight = FORM_HEIGHT
    End If
End Sub

Private Sub ApplyLocalizedCaptions()
    SetLocalizedCaption "lblDispatchLetters", "form.mail_dispatch.label.letters", "Письма"
    SetLocalizedCaption "lblDispatchSender", "form.mail_dispatch.label.sender", "Отправитель"
    SetLocalizedCaption "lblDispatchEnvelopeFormat", "form.mail_dispatch.label.envelope_format", "Формат конверта"
    SetLocalizedCaption "lblDispatchMailType", "form.mail_dispatch.label.mail_type", "Вид отправления"
    SetLocalizedCaption "lblDispatchMass", "form.mail_dispatch.label.mass", "Масса"
    SetLocalizedCaption "lblDispatchDeclaredValue", "form.mail_dispatch.label.declared_value", "Объявленная ценность"
    SetLocalizedCaption "lblDispatchComment", "form.mail_dispatch.label.comment", "Комментарий"
    SetLocalizedCaption "lblDispatchPreview", "form.mail_dispatch.label.preview", "Предпросмотр"

    btnDispatchRefresh.Caption = t("form.mail_dispatch.button.refresh", "Обновить")
    btnDispatchCreate.Caption = t("form.mail_dispatch.button.create", "Добавить в отправления")
    btnDispatchClose.Caption = t("form.mail_dispatch.button.close", "Закрыть")

    txtDispatchMailType.ControlTipText = t("form.mail_dispatch.tip.mail_type", "Например: заказное, простое, с уведомлением")
    txtDispatchMass.ControlTipText = t("form.mail_dispatch.tip.mass", "Масса отправления в граммах")
    txtDispatchDeclaredValue.ControlTipText = t("form.mail_dispatch.tip.declared_value", "Объявленная ценность в рублях")
    txtDispatchComment.ControlTipText = t("form.mail_dispatch.tip.comment", "Короткий служебный комментарий для отправления")
End Sub

Private Sub LoadDispatchData()
    LoadLettersList
    LoadSendersList
    LoadEnvelopeFormatList
End Sub

Private Sub LoadLettersList()
    Set allLettersData = RepositoryLoadLetterHistoryData()
    lstDispatchLetters.Clear

    Dim i As Long
    For i = 1 To allLettersData.count
        lstDispatchLetters.AddItem RepositoryFormatLetterHistoryDisplay(allLettersData(i))
    Next i
End Sub

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
    If lstDispatchLetters.ListCount > 0 Then
        lstDispatchLetters.ListIndex = 0
    End If

    If cmbDispatchEnvelopeFormat.ListCount > 0 Then
        cmbDispatchEnvelopeFormat.ListIndex = 0
    End If

    Dim defaultSender As String
    defaultSender = DispatchRepositoryGetDefaultSenderName()
    If Len(defaultSender) > 0 Then
        SelectComboValue cmbDispatchSender, defaultSender
    ElseIf cmbDispatchSender.ListCount > 0 Then
        cmbDispatchSender.ListIndex = 0
    End If

    If Len(Trim$(txtDispatchMailType.Text)) = 0 Then
        txtDispatchMailType.Text = t("form.mail_dispatch.default.mail_type", "заказное")
    End If
End Sub

Private Sub lstDispatchLetters_Click()
    UpdateDispatchPreview
End Sub

Private Sub btnDispatchRefresh_Click()
    LoadDispatchData
    SelectDefaultValues
    UpdateDispatchPreview
End Sub

Private Sub btnDispatchClose_Click()
    Unload Me
End Sub

Private Sub btnDispatchCreate_Click()
    On Error GoTo CreateError

    Dim record As clsLetterHistoryRecord
    Set record = GetSelectedHistoryRecord()
    If record Is Nothing Then
        MsgBox t("form.mail_dispatch.error.no_letter", "Выберите письмо для подготовки отправления."), vbExclamation
        Exit Sub
    End If

    If cmbDispatchSender.ListIndex < 0 Then
        MsgBox t("form.mail_dispatch.error.no_sender", "Выберите отправителя."), vbExclamation
        Exit Sub
    End If

    Dim envelopeFormatKey As String
    envelopeFormatKey = GetSelectedEnvelopeFormatKey()
    If Len(envelopeFormatKey) = 0 Then
        MsgBox t("form.mail_dispatch.error.no_envelope_format", "Выберите формат конверта."), vbExclamation
        Exit Sub
    End If

    Dim dispatchId As String
    dispatchId = DispatchRepositoryCreateItemFromHistoryRecord( _
        record, _
        cmbDispatchSender.Text, _
        envelopeFormatKey, _
        txtDispatchMailType.Text, _
        txtDispatchMass.Text, _
        txtDispatchDeclaredValue.Text, _
        txtDispatchComment.Text)

    If Len(dispatchId) = 0 Then
        MsgBox t("form.mail_dispatch.error.create_failed", "Не удалось добавить отправление в рабочую таблицу."), vbCritical
        Exit Sub
    End If

    MsgBox t("form.mail_dispatch.msg.created", "Отправление добавлено в таблицу почтовых отправлений.") & vbCrLf & dispatchId, vbInformation
    Exit Sub

CreateError:
    MsgBox t("form.mail_dispatch.error.runtime", "Ошибка подготовки почтового отправления: ") & Err.description, vbCritical
End Sub

Private Sub UpdateDispatchPreview()
    Dim record As clsLetterHistoryRecord
    Set record = GetSelectedHistoryRecord()

    If record Is Nothing Then
        txtDispatchPreview.Text = ""
        Exit Sub
    End If

    txtDispatchPreview.Text = DispatchRepositoryBuildRecipientPreviewByAddressee(record.Addressee)
End Sub

Private Function GetSelectedHistoryRecord() As clsLetterHistoryRecord
    If lstDispatchLetters.ListIndex < 0 Then Exit Function
    If allLettersData Is Nothing Then Exit Function
    If lstDispatchLetters.ListIndex + 1 > allLettersData.count Then Exit Function

    If TypeName(allLettersData(lstDispatchLetters.ListIndex + 1)) = "clsLetterHistoryRecord" Then
        Set GetSelectedHistoryRecord = allLettersData(lstDispatchLetters.ListIndex + 1)
    End If
End Function

Private Function GetSelectedEnvelopeFormatKey() As String
    If cmbDispatchEnvelopeFormat.ListIndex < 0 Then Exit Function
    If envelopeFormats Is Nothing Then Exit Function
    If cmbDispatchEnvelopeFormat.ListIndex + 1 > envelopeFormats.count Then Exit Function

    GetSelectedEnvelopeFormatKey = CStr(envelopeFormats(cmbDispatchEnvelopeFormat.ListIndex + 1)(EnvelopeFormatColumnKey))
End Function

Private Sub SelectComboValue(targetCombo As ComboBox, expectedValue As String)
    Dim i As Long
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(CStr(targetCombo.List(i)), expectedValue, vbTextCompare) = 0 Then
            targetCombo.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub SetLocalizedCaption(controlName As String, translationKey As String, fallbackText As String)
    On Error Resume Next
    Me.Controls(controlName).Caption = t(translationKey, fallbackText)
    On Error GoTo 0
End Sub
