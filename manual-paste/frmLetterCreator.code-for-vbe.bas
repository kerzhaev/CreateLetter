Option Explicit

'------------------------------------------------------------
'  ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
'------------------------------------------------------------
Private Const TOTAL_PAGES As Integer = 4
Public selectedAddressRow As Long
Private documentsList As Collection
Private contextMenuSelectedIndex As Integer

'------------------------------------------------------------
'  ИНИЦИАЛИЗАЦИЯ ФОРМЫ
'------------------------------------------------------------
Private Sub UserForm_Initialize()
    Set documentsList = New Collection
    contextMenuSelectedIndex = -1
    
    ConfigureFormAppearance
    InitializeControlValues
    LoadExecutorsList
    
    ConfigureMultilineTextBoxes
    ConfigureDocumentSumField
    
    ClearDocumentFields
    InitializeProgressInfo
    
    SwitchToPage 0
End Sub

Private Sub ConfigureDocumentSumField()
    On Error Resume Next
    
    If Not txtDocumentSum Is Nothing Then
        With txtDocumentSum
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .ControlTipText = T("form.letter_creator.tip.document_sum", "Сумма документа в рублях (необязательно). Например: 125000")
            .Value = ""
            .backColor = RGB(255, 255, 255)
        End With
        Debug.Print "Поле суммы документа настроено"
    End If
    
    On Error GoTo 0
End Sub

Private Sub InitializeProgressInfo()
    lblProgressInfo.Caption = T("form.letter_creator.progress.page", "Шаг") & " 1 из " & TOTAL_PAGES
    lblAttachmentsCount.Caption = T("form.letter_creator.attachments_count", "Выбрано документов:") & " 0"
End Sub

'------------------------------------------------------------
'  ИНИЦИАЛИЗАЦИЯ ЗНАЧЕНИЙ ЭЛЕМЕНТОВ УПРАВЛЕНИЯ
'------------------------------------------------------------
Private Sub InitializeControlValues()
    On Error Resume Next
    
    Me.Controls("txtLetterDate").Value = Format(Date, "dd.mm.yyyy")
    Me.Controls("txtLetterNumber").Value = "7/"
    
    With Me.Controls("cmbDocumentType")
        .Clear
        .AddItem "Чужие подтверждённые документы"
        .AddItem "Свои для подтверждения"
        .ListIndex = 0
    End With
    
    With Me.Controls("cmbLetterType")
        .Clear
        .AddItem "Обычное"
        .AddItem "ДСП"
        .ListIndex = 0
    End With
    
    selectedAddressRow = 0
    On Error GoTo 0
End Sub

'------------------------------------------------------------
'  СОБЫТИЯ ИЗМЕНЕНИЯ ПОЛЕЙ
'------------------------------------------------------------
Private Sub txtCity_Change()
    AutoResizeTextBoxHeight "txtCity"
    CheckRequiredFields
End Sub

Private Sub txtRegion_Change()
    AutoResizeTextBoxHeight "txtRegion"
    CheckRequiredFields
End Sub

Private Sub txtPostalCode_Change()
    AutoResizeTextBoxHeight "txtPostalCode"
    CheckRequiredFields
End Sub

Private Sub cmbExecutor_Change()
    CheckRequiredFields
End Sub

Private Sub txtAddressee_Change()
    AutoResizeTextBoxHeight "txtAddressee"
    CheckRequiredFields
End Sub

Private Sub txtStreet_Change()
    AutoResizeTextBoxHeight "txtStreet"
End Sub

Private Sub txtDistrict_Change()
    AutoResizeTextBoxHeight "txtDistrict"
End Sub

Private Sub txtDocumentSum_Change()
    On Error Resume Next
    
    If Not txtDocumentSum Is Nothing Then
        Dim currentValue As String
        currentValue = Trim(txtDocumentSum.Value)
        
        If Len(currentValue) > 0 And Not IsNumeric(currentValue) Then
            txtDocumentSum.backColor = RGB(255, 240, 240)
        Else
            txtDocumentSum.backColor = RGB(255, 255, 255)
        End If
    End If
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------
'  СОБЫТИЯ ПОТЕРИ ФОКУСА
'------------------------------------------------------------
Private Sub txtCity_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

Private Sub txtRegion_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

Private Sub txtPostalCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

Private Sub txtAddresseePhone_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

Private Sub txtAddressee_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

Private Sub txtStreet_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

Private Sub txtDistrict_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If selectedAddressRow > 1 Then
        AutoUpdateAddressIfChanged
    End If
End Sub

'------------------------------------------------------------
'  КНОПКА ИСТОРИИ ПИСЕМ
'------------------------------------------------------------
Private Sub btnLetterHistory_Click()
    On Error GoTo HistoryError
    
    Me.Hide
    
    Dim existingForm As Object
    On Error Resume Next
    Set existingForm = VBA.UserForms("frmLetterHistory")
    On Error GoTo HistoryError
    
    If Not existingForm Is Nothing Then
        existingForm.SetFocus
        existingForm.ZOrder 0
        Debug.Print "Форма истории уже открыта - активирована"
    Else
        Load frmLetterHistory
        frmLetterHistory.Show vbModeless
        Debug.Print "Форма истории открыта немодально"
    End If
    
    Exit Sub
    
HistoryError:
    MsgBox "Ошибка при открытии формы истории: " & Err.description, vbCritical
End Sub

'------------------------------------------------------------
'  КНОПКА ОЧИСТКИ ПОИСКА
'------------------------------------------------------------
Private Sub btnClearSearch_Click()
    On Error Resume Next
    
    Me.Controls("txtAddressSearch").Value = ""
    Me.Controls("lstAddresses").Clear
    
    ClearAllAddressFields
    ResetAddressFormState
    
    Me.Controls("txtAddressSearch").SetFocus
    
    Debug.Print "Выполнена полная очистка поиска и полей адреса"
    
    On Error GoTo 0
End Sub

Private Sub ClearAllAddressFields()
    On Error Resume Next
    
    Dim addressFields As Variant
    Dim i As Long
    Dim ctrl As Control
    
    addressFields = Array("txtAddressee", "txtStreet", "txtCity", "txtDistrict", "txtRegion", "txtPostalCode", "txtAddresseePhone")
    
    For i = LBound(addressFields) To UBound(addressFields)
        Set ctrl = Me.Controls(addressFields(i))
        If Not ctrl Is Nothing Then
            ctrl.Value = ""
            ctrl.backColor = RGB(255, 255, 255)
            Debug.Print "Очищено поле: " & addressFields(i)
        End If
    Next i
    
    selectedAddressRow = 0
    Debug.Print "Все поля адреса очищены"
    
    On Error GoTo 0
End Sub

Private Sub CheckRequiredFields()
    On Error Resume Next
    
    If Len(Trim(Me.Controls("txtCity").Value)) = 0 Then
        Me.Controls("txtCity").backColor = RGB(255, 240, 240)
    Else
        Me.Controls("txtCity").backColor = RGB(240, 255, 240)
    End If
    
    If Len(Trim(Me.Controls("txtRegion").Value)) = 0 Then
        Me.Controls("txtRegion").backColor = RGB(255, 240, 240)
    Else
        Me.Controls("txtRegion").backColor = RGB(240, 255, 240)
    End If
    
    If Len(Trim(Me.Controls("txtPostalCode").Value)) = 0 Then
        Me.Controls("txtPostalCode").backColor = RGB(255, 240, 240)
    Else
        Me.Controls("txtPostalCode").backColor = RGB(240, 255, 240)
    End If
    
    If Me.Controls("cmbExecutor").ListIndex < 0 Or Len(Trim(Me.Controls("cmbExecutor").Value)) = 0 Then
        Me.Controls("cmbExecutor").backColor = RGB(255, 240, 240)
    Else
        Me.Controls("cmbExecutor").backColor = RGB(240, 255, 240)
    End If
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------
'  НАСТРОЙКА ВНЕШНЕГО ВИДА
'------------------------------------------------------------
Private Sub ConfigureFormAppearance()
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 10
    Me.Caption = T("form.letter_creator.title", "Формирование писем") & " v1.6.0"
    
    On Error Resume Next
    
    If Not lstSelectedAttachments Is Nothing Then
        lstSelectedAttachments.Font.Size = 9
        lstSelectedAttachments.ControlTipText = T("form.letter_creator.tip.selected_attachments", "Для просмотра полного названия наведите на элемент")
        lstSelectedAttachments.IntegralHeight = False
    End If
    
    If Not btnEditAddress Is Nothing Then
        btnEditAddress.Caption = T("form.letter_creator.caption.edit_address", "Изменить адрес")
        btnEditAddress.ControlTipText = T("form.letter_creator.tip.edit_address", "Редактировать выбранный адрес")
        btnEditAddress.Enabled = False
    End If
    
    If Not btnDeleteAddress Is Nothing Then
        btnDeleteAddress.Caption = T("form.letter_creator.caption.delete_address", "Удалить адрес")
        btnDeleteAddress.ControlTipText = T("form.letter_creator.tip.delete_address", "Удалить выбранный адрес")
        btnDeleteAddress.Enabled = False
    End If
    
    If Not txtAddresseePhone Is Nothing Then
        txtAddresseePhone.ControlTipText = T("form.letter_creator.tip.phone", "Телефон адресата (формат: 8-xxx-xxx-xx-xx)")
        txtAddresseePhone.Enabled = True
        txtAddresseePhone.backColor = RGB(255, 255, 255)
    End If
    
    If Not btnLetterHistory Is Nothing Then
        With btnLetterHistory
            .Caption = T("form.letter_creator.caption.letter_history", "История писем")
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .backColor = RGB(156, 39, 176)
            .ForeColor = RGB(255, 255, 255)
            .ControlTipText = T("form.letter_creator.tip.letter_history", "Открыть форму истории отправленных писем")
        End With
    End If
    
    On Error GoTo 0
    
    txtAddressSearch.ControlTipText = T("form.letter_creator.tip.address_search", "Введите часть наименования для поиска адресата")
    txtLetterNumber.ControlTipText = T("form.letter_creator.tip.letter_number", "Введите номер после 7/ (например: 125 > получится 7/125)")
    txtLetterDate.ControlTipText = T("form.letter_creator.tip.letter_date", "Формат: дд.мм.гггг")
End Sub

Private Sub txtAddresseePhone_Change()
    On Error Resume Next
    
    Dim currentValue As String, cursorPos As Long
    currentValue = Me.Controls("txtAddresseePhone").Value
    cursorPos = Me.Controls("txtAddresseePhone").SelStart
    
    If Len(currentValue) >= 7 Then
        Dim formattedPhone As String
        formattedPhone = FormatPhoneNumber(currentValue)
        
        If formattedPhone <> currentValue Then
            Me.Controls("txtAddresseePhone").Value = formattedPhone
            Me.Controls("txtAddresseePhone").SelStart = WorksheetFunction.Min(cursorPos + (Len(formattedPhone) - Len(currentValue)), Len(formattedPhone))
        End If
    End If
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------
'  ИСПОЛНИТЕЛИ
'------------------------------------------------------------
Private Sub LoadExecutorsList()
    On Error Resume Next
    Dim col As Collection, i As Long
    Set col = GetExecutorsList()
    
    If Not cmbExecutor Is Nothing Then
        cmbExecutor.Clear
        For i = 1 To col.count
            cmbExecutor.AddItem col(i)
        Next i
    End If
    On Error GoTo 0
End Sub

'------------------------------------------------------------
'  ОЧИСТКА ПОЛЕЙ ДОКУМЕНТА
'------------------------------------------------------------
Private Sub ClearDocumentFields()
    On Error Resume Next
    If Not txtDocNumber Is Nothing Then txtDocNumber.Value = ""
    If Not txtDocDate Is Nothing Then txtDocDate.Value = ""
    If Not txtDocCopies Is Nothing Then txtDocCopies.Value = ""
    If Not txtDocSheets Is Nothing Then txtDocSheets.Value = ""
    If Not txtDocumentSum Is Nothing Then txtDocumentSum.Value = ""
    On Error GoTo 0
End Sub

'=====================================================================
'                       НАВИГАЦИЯ
'=====================================================================
Private Sub btnPrevious_Click()
    If mpgWizard.Value > 0 Then SwitchToPage mpgWizard.Value - 1
End Sub

Private Sub btnNext_Click()
    Dim cur As Integer: cur = mpgWizard.Value
    
    If Not ValidatePage(cur) Then Exit Sub
    
    If cur = TOTAL_PAGES - 1 Then
        If ValidateForm Then
            UpdateSummaryInfo
            CreateWordLetter
            SaveLetterToDatabase
            MsgBox "Письмо успешно создано!", vbInformation
            Unload Me
        End If
    Else
        If cur = 2 Then UpdateSummaryInfo
        SwitchToPage cur + 1
    End If
End Sub

Private Sub SwitchToPage(pg As Integer)
    If pg < 0 Or pg > TOTAL_PAGES - 1 Then Exit Sub
    
    mpgWizard.Value = pg
    lblProgressInfo.Caption = "Шаг " & pg + 1 & " из " & TOTAL_PAGES
    
    btnPrevious.Enabled = (pg > 0)
    
    If pg = TOTAL_PAGES - 1 Then
        btnNext.Caption = "СОЗДАТЬ ПИСЬМО"
        btnNext.backColor = RGB(76, 175, 80)
        btnNext.ForeColor = RGB(255, 255, 255)
        btnNext.Font.Bold = True
        btnNext.Font.Size = 11
    Else
        btnNext.Caption = "Далее >"
        btnNext.backColor = RGB(240, 240, 240)
        btnNext.ForeColor = RGB(0, 0, 0)
        btnNext.Font.Bold = False
        btnNext.Font.Size = 10
    End If
    
    SetFocusToFirstControl pg
End Sub

Private Sub SetFocusToFirstControl(pg As Integer)
    Select Case pg
        Case 0: SafeSetFocus "txtAddressSearch"
        Case 1: SafeSetFocus "txtLetterNumber"
        Case 2: SafeSetFocus "txtAttachmentSearch"
        Case 3: SafeSetFocus "btnNext"
    End Select
End Sub

'=====================================================================
'                 ВАЛИДАЦИЯ ШАГОВ
'=====================================================================
Private Function ValidatePage(pg As Integer) As Boolean
    ValidatePage = False
    
    Select Case pg
        Case 0
            On Error Resume Next
            
            If Trim(Me.Controls("txtAddressee").Value) = "" Then
                MsgBox "Заполните поле 'Наименование получателя'", vbExclamation
                Me.Controls("txtAddressee").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            If Trim(Me.Controls("txtCity").Value) = "" Then
                MsgBox "Заполните поле 'Город' - это обязательное поле", vbExclamation
                Me.Controls("txtCity").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            If Trim(Me.Controls("txtRegion").Value) = "" Then
                MsgBox "Заполните поле 'Регион' - это обязательное поле", vbExclamation
                Me.Controls("txtRegion").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            If Trim(Me.Controls("txtPostalCode").Value) = "" Then
                MsgBox "Заполните поле 'Почтовый индекс' - это обязательное поле", vbExclamation
                Me.Controls("txtPostalCode").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            If Len(Trim(Me.Controls("txtAddresseePhone").Value)) > 0 Then
                If Not IsPhoneNumberValid(Me.Controls("txtAddresseePhone").Value) Then
                    MsgBox "Введите корректный номер телефона адресата", vbExclamation
                    Me.Controls("txtAddresseePhone").SetFocus
                    On Error GoTo 0: Exit Function
                End If
            End If
            
            On Error GoTo 0
            ValidateAndUpdateSelectedAddress
            
        Case 1
            On Error Resume Next
            
            If Trim(Me.Controls("txtLetterNumber").Value) = "" Then
                MsgBox "Введите номер исходящего письма.", vbExclamation
                Me.Controls("txtLetterNumber").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            If Trim(Me.Controls("txtLetterDate").Value) = "" Then
                MsgBox "Введите дату письма.", vbExclamation
                Me.Controls("txtLetterDate").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            If Me.Controls("cmbExecutor").ListIndex < 0 Or Len(Trim(Me.Controls("cmbExecutor").Value)) = 0 Then
                MsgBox "Выберите исполнителя - это обязательное поле!", vbExclamation
                Me.Controls("cmbExecutor").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            Dim d As Date
            If Not TryParseDate(Me.Controls("txtLetterDate").Value, d) Then
                MsgBox "Некорректный формат даты письма.", vbExclamation
                Me.Controls("txtLetterDate").SetFocus
                On Error GoTo 0: Exit Function
            End If
            
            On Error GoTo 0
            
        Case 2
            If documentsList.count = 0 Then
                MsgBox "Добавьте хотя бы один документ-приложение.", vbExclamation
                Exit Function
            End If
    End Select
    
    ValidatePage = True
End Function

'=====================================================================
'                 ЗАЩИТА ПРЕФИКСА В НОМЕРЕ ПИСЬМА
'=====================================================================
Private Sub txtLetterNumber_Change()
    On Error Resume Next
    If Not txtLetterNumber Is Nothing Then
        Dim currentValue As String
        currentValue = txtLetterNumber.Value
        
        If Left(currentValue, 2) <> "7/" Then
            Dim numericPart As String
            numericPart = Replace(currentValue, "7/", "")
            
            txtLetterNumber.Value = "7/" & numericPart
            txtLetterNumber.SelStart = Len(txtLetterNumber.Value)
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub txtLetterNumber_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    If Not txtLetterNumber Is Nothing Then
        If txtLetterNumber.SelStart < 2 And (KeyCode = 8 Or KeyCode = 46) Then
            txtLetterNumber.SelStart = 2
            KeyCode = 0
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub txtLetterNumber_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    If Not txtLetterNumber Is Nothing Then
        If txtLetterNumber.SelStart < 2 Then
            txtLetterNumber.SelStart = 2
        End If
    End If
    On Error GoTo 0
End Sub

'=====================================================================
'      ШАГ 1 - поиск и выбор адресата
'=====================================================================
Private Sub txtAddressSearch_Change()
    On Error Resume Next
    
    If Not Me.Controls("lstAddresses") Is Nothing Then
        Me.Controls("lstAddresses").Clear
        
        ResetAddressFormState
        
        If Len(Trim(Me.Controls("txtAddressSearch").Value)) > 0 Then
            Dim res As Collection, i As Long
            Set res = GetCachedAddresses(Me.Controls("txtAddressSearch").Value)
            For i = 1 To res.count
                Me.Controls("lstAddresses").AddItem res(i)
            Next i
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub ResetAddressFormState()
    On Error Resume Next
    
    If selectedAddressRow > 1 Then
        ValidateAndUpdateSelectedAddress
    End If
    
    selectedAddressRow = 0
    
    If Not Me.Controls("btnSaveNewAddress") Is Nothing Then
        Me.Controls("btnSaveNewAddress").Enabled = True
    End If
    
    If Not Me.Controls("btnEditAddress") Is Nothing Then
        Me.Controls("btnEditAddress").Enabled = False
    End If
    
    If Not Me.Controls("btnDeleteAddress") Is Nothing Then
        Me.Controls("btnDeleteAddress").Enabled = False
    End If
    
    Debug.Print "Состояние формы адреса сброшено"
    
    On Error GoTo 0
End Sub

Private Sub lstAddresses_Click()
    On Error Resume Next
    
    If lstAddresses Is Nothing Or lstAddresses.ListIndex < 0 Then Exit Sub
    
    Dim itm As String, parts As Variant
    itm = lstAddresses.List(lstAddresses.ListIndex)
    
    If InStr(itm, " | ") = 0 Then
        MsgBox "Неверный формат записи адреса.", vbExclamation
        Exit Sub
    End If
    
    parts = Split(itm, " | ")
    If UBound(parts) < 7 Then
        MsgBox "Данных адреса недостаточно.", vbExclamation
        Exit Sub
    End If
    
    Me.Controls("txtAddressee").Value = parts(0)
    Me.Controls("txtStreet").Value = parts(1)
    Me.Controls("txtCity").Value = parts(2)
    Me.Controls("txtDistrict").Value = parts(3)
    Me.Controls("txtRegion").Value = parts(4)
    Me.Controls("txtPostalCode").Value = parts(5)
    Me.Controls("txtAddresseePhone").Value = parts(6)
    
    selectedAddressRow = CLng(parts(7))
    
    If Not btnSaveNewAddress Is Nothing Then btnSaveNewAddress.Enabled = False
    
    If Not btnEditAddress Is Nothing Then btnEditAddress.Enabled = True
    If Not btnDeleteAddress Is Nothing Then btnDeleteAddress.Enabled = True
    
    On Error GoTo 0
End Sub

Private Sub btnSaveNewAddress_Click()
    On Error Resume Next
    If txtAddressee Is Nothing Or Trim(txtAddressee.Value) = "" Then
        MsgBox "Введите наименование адресата!", vbExclamation
        Exit Sub
    End If
    
    If IsAddressDuplicate(CreateAddressArray) Then
        MsgBox "Такой адрес уже существует!", vbExclamation
        Exit Sub
    End If
    
    SaveNewAddress CreateAddressArray
    MsgBox "Адрес сохранён!", vbInformation
    
    ClearAddressCache
    
    On Error GoTo 0
End Sub

'=====================================================================
'      КНОПКИ РЕДАКТИРОВАНИЯ АДРЕСА
'=====================================================================
Private Sub btnEditAddress_Click()
    On Error Resume Next
    
    If selectedAddressRow <= 1 Then
        MsgBox "Выберите адрес для редактирования!", vbExclamation
        Exit Sub
    End If
    
    ValidateAndUpdateSelectedAddress
    
    Dim addressArray As Variant
    addressArray = CreateAddressArray()
    
    If IsAddressDuplicate(addressArray, selectedAddressRow) Then
        MsgBox "Адрес с такими данными уже существует!", vbExclamation
        Exit Sub
    End If
    
    UpdateExistingAddress selectedAddressRow, addressArray
    
    ClearAddressCache
    txtAddressSearch_Change
    
    MsgBox "Адрес успешно обновлен!", vbInformation
    On Error GoTo 0
End Sub

Private Sub btnDeleteAddress_Click()
    On Error GoTo DeleteError
    
    If selectedAddressRow = 0 Then
        MsgBox "Выберите адрес для удаления!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Вы уверены, что хотите удалить этот адрес?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        DeleteExistingAddress selectedAddressRow
        MsgBox "Адрес успешно удалён!", vbInformation
        
        ClearAddressFields
        ClearAddressCache
        
        selectedAddressRow = 0
        btnEditAddress.Enabled = False
        btnDeleteAddress.Enabled = False
    End If
    
    Exit Sub
    
DeleteError:
    MsgBox "Ошибка при удалении адреса: " & Err.description, vbCritical
End Sub

Private Sub ClearAddressFields()
    On Error Resume Next
    If Not txtAddressee Is Nothing Then txtAddressee.Value = ""
    If Not txtStreet Is Nothing Then txtStreet.Value = ""
    If Not txtCity Is Nothing Then txtCity.Value = ""
    If Not txtDistrict Is Nothing Then txtDistrict.Value = ""
    If Not txtRegion Is Nothing Then txtRegion.Value = ""
    If Not txtPostalCode Is Nothing Then txtPostalCode.Value = ""
    On Error GoTo 0
End Sub

'=====================================================================
'      ШАГ 3 - добавление приложений
'=====================================================================
Private Sub txtAttachmentSearch_Change()
    On Error Resume Next
    If Not lstAvailableAttachments Is Nothing Then
        lstAvailableAttachments.Clear
        If Len(Trim(txtAttachmentSearch.Value)) > 0 Then
            Dim res As Collection, i As Long
            Set res = GetCachedAttachments(txtAttachmentSearch.Value)
            For i = 1 To res.count: lstAvailableAttachments.AddItem res(i): Next i
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub lstAvailableAttachments_Click()
    On Error Resume Next
    If Not txtDocNumber Is Nothing Then
        txtDocNumber.SetFocus
    End If
    On Error GoTo 0
End Sub

Private Sub lstAvailableAttachments_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    If lstAvailableAttachments.ListIndex >= 0 And Not txtDocNumber Is Nothing Then
        txtDocNumber.SetFocus
    End If
    On Error GoTo 0
End Sub

'=====================================================================
'      ДОБАВЛЕНИЕ ПРИЛОЖЕНИЙ С СУММОЙ
'=====================================================================
Private Sub btnAddAttachment_Click()
    On Error Resume Next
    
    If lstAvailableAttachments Is Nothing Or lstAvailableAttachments.ListIndex < 0 Then
        MsgBox "Выберите документ в левом списке!", vbExclamation
        Exit Sub
    End If
    
    Dim docArr As Variant
    docArr = CreateDocumentArrayWithSum( _
        lstAvailableAttachments.List(lstAvailableAttachments.ListIndex), _
        Trim(IIf(txtDocNumber Is Nothing, "", txtDocNumber.Value)), _
        Trim(IIf(txtDocDate Is Nothing, "", txtDocDate.Value)), _
        Trim(IIf(txtDocCopies Is Nothing, "", txtDocCopies.Value)), _
        Trim(IIf(txtDocSheets Is Nothing, "", txtDocSheets.Value)), _
        Trim(IIf(txtDocumentSum Is Nothing, "", txtDocumentSum.Value)))
    
    documentsList.Add docArr
    
    If Not lstSelectedAttachments Is Nothing Then
        lstSelectedAttachments.AddItem FormatDocumentNameWithSum(docArr)
    End If
    
    If Not lblAttachmentsCount Is Nothing Then
        lblAttachmentsCount.Caption = "Выбрано документов: " & documentsList.count
    End If
    
    ClearDocumentFields
    On Error GoTo 0
End Sub

Private Sub btnRemoveAttachment_Click()
    On Error Resume Next
    
    If lstSelectedAttachments Is Nothing Or lstSelectedAttachments.ListIndex < 0 Then
        MsgBox "Выберите документ в правом списке!", vbExclamation
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstSelectedAttachments.ListIndex
    
    If selectedIndex + 1 <= documentsList.count Then
        documentsList.Remove selectedIndex + 1
        
        lstSelectedAttachments.Clear
        Dim i As Long
        For i = 1 To documentsList.count
            lstSelectedAttachments.AddItem FormatDocumentNameWithSum(documentsList(i))
        Next i
    End If
    
    If Not lblAttachmentsCount Is Nothing Then
        lblAttachmentsCount.Caption = "Выбрано документов: " & documentsList.count
    End If
    
    On Error GoTo 0
End Sub

'=====================================================================
'                    КОНТЕКСТНОЕ МЕНЮ
'=====================================================================
Private Sub lstSelectedAttachments_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 And lstSelectedAttachments.ListIndex >= 0 Then
        contextMenuSelectedIndex = lstSelectedAttachments.ListIndex
        ShowSimpleContextMenu
    End If
End Sub

Private Sub ShowSimpleContextMenu()
    On Error GoTo MenuError
    
    Dim menuChoice As String
    menuChoice = InputBox("Выберите действие:" & vbCrLf & _
                         "1 - Редактировать реквизиты" & vbCrLf & _
                         "2 - Дублировать документ" & vbCrLf & _
                         "3 - Удалить из списка" & vbCrLf & _
                         "4 - Переместить вверх" & vbCrLf & _
                         "5 - Переместить вниз", _
                         "Действия с документом", "1")
    
    Select Case menuChoice
        Case "1": EditDocumentRequisites
        Case "2": DuplicateDocument
        Case "3": RemoveSelectedDocument
        Case "4": If contextMenuSelectedIndex > 0 Then MoveDocumentUp
        Case "5": If contextMenuSelectedIndex < lstSelectedAttachments.ListCount - 1 Then MoveDocumentDown
        Case Else: Exit Sub
    End Select
    
    Exit Sub
    
MenuError:
End Sub

Public Sub EditDocumentRequisites()
    On Error GoTo EditError
    
    If contextMenuSelectedIndex >= 0 And contextMenuSelectedIndex < documentsList.count Then
        Dim docArray As Variant
        docArray = documentsList.item(contextMenuSelectedIndex + 1)
        
        If IsArray(docArray) Then
            If UBound(docArray) >= 4 Then
                txtDocNumber.Value = docArray(1)
                txtDocDate.Value = docArray(2)
                txtDocCopies.Value = docArray(3)
                txtDocSheets.Value = docArray(4)
            End If
            
            If UBound(docArray) >= 5 And Not txtDocumentSum Is Nothing Then
                txtDocumentSum.Value = docArray(5)
            End If
        End If
    End If
    Exit Sub
    
EditError:
End Sub

Public Sub DuplicateDocument()
    On Error GoTo DuplicateError
    
    If contextMenuSelectedIndex >= 0 And contextMenuSelectedIndex < documentsList.count Then
        Dim docIndex As Long
        docIndex = contextMenuSelectedIndex + 1
        
        Dim sourceName As String, sourceDate As String, sourceCopies As String, sourceSheets As String, sourceSum As String
        sourceName = ""
        sourceDate = ""
        sourceCopies = ""
        sourceSheets = ""
        sourceSum = ""
        
        Dim sourceItem As Variant
        sourceItem = documentsList.item(docIndex)
        
        If IsArray(sourceItem) Then
            If UBound(sourceItem) >= 4 Then
                sourceName = CStr(sourceItem(0))
                sourceDate = CStr(sourceItem(2))
                sourceCopies = CStr(sourceItem(3))
                sourceSheets = CStr(sourceItem(4))
            End If
            
            If UBound(sourceItem) >= 5 Then
                sourceSum = CStr(sourceItem(5))
            End If
        End If
        
        Dim duplicateDoc As Variant
        duplicateDoc = CreateDocumentArrayWithSum(sourceName, "", sourceDate, sourceCopies, sourceSheets, sourceSum)
        
        documentsList.Add duplicateDoc
        lstSelectedAttachments.AddItem FormatDocumentNameWithSum(duplicateDoc)
        
        If Not lblAttachmentsCount Is Nothing Then
            lblAttachmentsCount.Caption = "Выбрано документов: " & documentsList.count
        End If
    End If
    Exit Sub
    
DuplicateError:
    MsgBox "Ошибка при дублировании документа: " & Err.description, vbCritical
End Sub

Public Sub RemoveSelectedDocument()
    On Error GoTo RemoveError
    
    If contextMenuSelectedIndex >= 0 And contextMenuSelectedIndex < documentsList.count Then
        documentsList.Remove contextMenuSelectedIndex + 1
        lstSelectedAttachments.RemoveItem contextMenuSelectedIndex
        
        If Not lblAttachmentsCount Is Nothing Then
            lblAttachmentsCount.Caption = "Выбрано документов: " & documentsList.count
        End If
    End If
    Exit Sub
    
RemoveError:
End Sub

Public Sub MoveDocumentUp()
    On Error GoTo MoveUpError
    
    If contextMenuSelectedIndex > 0 Then
        Dim tempDoc As Variant
        tempDoc = documentsList(contextMenuSelectedIndex)
        documentsList.Remove contextMenuSelectedIndex
        documentsList.Add tempDoc, , contextMenuSelectedIndex
        
        RefreshDocumentsList
        lstSelectedAttachments.ListIndex = contextMenuSelectedIndex - 1
    End If
    Exit Sub
    
MoveUpError:
End Sub

Public Sub MoveDocumentDown()
    On Error GoTo MoveDownError
    
    If contextMenuSelectedIndex < documentsList.count - 1 Then
        Dim tempDoc As Variant
        tempDoc = documentsList(contextMenuSelectedIndex + 2)
        documentsList.Remove contextMenuSelectedIndex + 2
        documentsList.Add tempDoc, , contextMenuSelectedIndex + 1
        
        RefreshDocumentsList
        lstSelectedAttachments.ListIndex = contextMenuSelectedIndex + 1
    End If
    Exit Sub
    
MoveDownError:
End Sub

Private Sub RefreshDocumentsList()
    lstSelectedAttachments.Clear
    Dim i As Long
    For i = 1 To documentsList.count
        lstSelectedAttachments.AddItem FormatDocumentNameWithSum(documentsList(i))
    Next i
End Sub

'=====================================================================
'      ШАГ 4 - сводка и создание письма
'=====================================================================
Private Sub UpdateSummaryInfo()
    On Error Resume Next
    
    If Not lblSummaryRecipient Is Nothing Then
        lblSummaryRecipient.Caption = IIf(txtAddressee Is Nothing, "", txtAddressee.Value)
    End If
    
    If Not lblSummaryNumber Is Nothing Then
        lblSummaryNumber.Caption = IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value)
    End If
    
    If Not lblSummaryDate Is Nothing Then
        lblSummaryDate.Caption = IIf(txtLetterDate Is Nothing, "", txtLetterDate.Value)
    End If
    
    If Not lblSummaryExecutor Is Nothing Then
        lblSummaryExecutor.Caption = IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value)
    End If
    
    If Not lblSummaryDocsCount Is Nothing Then
        lblSummaryDocsCount.Caption = CStr(documentsList.count)
    End If
    
    If Not txtFinalAttachments Is Nothing Then
        Dim attachmentText As String
        attachmentText = ""
        
        Dim i As Long
        For i = 1 To documentsList.count
            If i > 1 Then attachmentText = attachmentText & vbCrLf
            attachmentText = attachmentText & i & ". " & FormatDocumentNameWithSum(documentsList(i)) & ";"
        Next i
        
        txtFinalAttachments.Value = attachmentText
    End If
    
    On Error GoTo 0
End Sub

'=====================================================================
'  ГЛОБАЛЬНАЯ ПРОВЕРКА ПЕРЕД СОЗДАНИЕМ
'=====================================================================
Private Function ValidateForm() As Boolean
    ValidateForm = False
    
    On Error Resume Next
    
    If txtAddressee Is Nothing Or Trim(txtAddressee.Value) = "" Then
        MsgBox "Адресат не заполнен!", vbExclamation: SwitchToPage 0: Exit Function
    End If
    
    If txtCity Is Nothing Or Trim(txtCity.Value) = "" Then
        MsgBox "Город не заполнен!", vbExclamation: SwitchToPage 0: Exit Function
    End If
    
    If txtRegion Is Nothing Or Trim(txtRegion.Value) = "" Then
        MsgBox "Регион не заполнен!", vbExclamation: SwitchToPage 0: Exit Function
    End If
    
    If txtPostalCode Is Nothing Or Trim(txtPostalCode.Value) = "" Then
        MsgBox "Почтовый индекс не заполнен!", vbExclamation: SwitchToPage 0: Exit Function
    End If
    
    If txtLetterNumber Is Nothing Or Trim(txtLetterNumber.Value) = "" Then
        MsgBox "Номер письма не заполнен!", vbExclamation: SwitchToPage 1: Exit Function
    End If
    
    If txtLetterDate Is Nothing Or Trim(txtLetterDate.Value) = "" Then
        MsgBox "Дата письма не заполнена!", vbExclamation: SwitchToPage 1: Exit Function
    End If
    
    If cmbExecutor Is Nothing Or cmbExecutor.ListIndex < 0 Or Len(Trim(cmbExecutor.Value)) = 0 Then
        MsgBox "Исполнитель не выбран!", vbExclamation: SwitchToPage 1: Exit Function
    End If
    
    If documentsList.count = 0 Then
        MsgBox "Добавьте хотя бы один документ!", vbExclamation: SwitchToPage 2: Exit Function
    End If
    
    On Error GoTo 0
    
    ValidateForm = True
End Function

'=====================================================================
'  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ АДРЕСА
'=====================================================================
Private Function CreateAddressArray() As Variant
    Dim arr(6) As String
    
    On Error Resume Next
    arr(0) = Me.Controls("txtAddressee").Value
    arr(1) = Me.Controls("txtStreet").Value
    arr(2) = Me.Controls("txtCity").Value
    arr(3) = Me.Controls("txtDistrict").Value
    arr(4) = Me.Controls("txtRegion").Value
    arr(5) = Me.Controls("txtPostalCode").Value
    arr(6) = Me.Controls("txtAddresseePhone").Value
    On Error GoTo 0
    
    CreateAddressArray = arr
End Function

'=====================================================================
'  НОВЫЕ ФУНКЦИИ ДЛЯ РАБОТЫ С СУММОЙ ДОКУМЕНТОВ
'=====================================================================
Public Function CreateDocumentArrayWithSum(docName As String, docNumber As String, docDate As String, docCopies As String, docSheets As String, docSum As String) As Variant
    Dim docArray(5) As String
    docArray(0) = Trim(docName)
    docArray(1) = Trim(docNumber)
    docArray(2) = Trim(docDate)
    docArray(3) = Trim(docCopies)
    docArray(4) = Trim(docSheets)
    docArray(5) = Trim(docSum)
    
    CreateDocumentArrayWithSum = docArray
End Function

Public Function FormatDocumentNameWithSum(docArray As Variant) As String
    If Not IsArray(docArray) Then
        FormatDocumentNameWithSum = "Ошибка: неверный формат данных"
        Exit Function
    End If
    
    Dim result As String
    result = docArray(0)
    
    result = result & " №"
    If Len(Trim(docArray(1))) > 0 Then
        result = result & docArray(1)
    Else
        result = result & "    "
    End If
    
    result = result & " от "
    If Len(Trim(docArray(2))) > 0 Then
        result = result & docArray(2)
    Else
        result = result & "        "
    End If
    
    ' ИСПРАВЛЕНО: Проверяем размер массива перед обращением к элементу суммы
    If UBound(docArray) >= 5 And Len(Trim(docArray(5))) > 0 Then
        If IsNumeric(docArray(5)) Then
            ' ИСПРАВЛЕНО: Убираем разделители тысяч для предотвращения неразрывных пробелов
            result = result & " на сумму " & CStr(CLng(CDbl(docArray(5)))) & " руб."
        Else
            result = result & " (" & docArray(5) & ")"
        End If
    End If
    
    result = result & " ("
    
    If Len(Trim(docArray(3))) > 0 Then
        result = result & docArray(3) & " экз."
    Else
        result = result & "  экз."
    End If
    
    result = result & ", "
    If Len(Trim(docArray(4))) > 0 Then
        result = result & docArray(4) & " л."
    Else
        result = result & "   л."
    End If
    
    result = result & ")"
    
    FormatDocumentNameWithSum = result
End Function


'=====================================================================
'      СОЗДАНИЕ ПИСЬМА В WORD
'=====================================================================
Private Sub CreateWordLetter()
    Dim wordApp As Object
    Dim wordDoc As Object
    
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo ErrorHandler
    
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    
    If wordApp Is Nothing Then
        Err.Raise 429, "CreateWordLetter", "Не удалось создать объект Word.Application"
    End If
    
    wordApp.Visible = True
    
    Dim templatePath As String
    If Not cmbLetterType Is Nothing And cmbLetterType.ListIndex = 1 Then
        templatePath = ThisWorkbook.Path & "\ШаблонПисьмаДСП.docx"
    Else
        templatePath = ThisWorkbook.Path & "\ШаблонПисьма.docx"
    End If
    
    If dir(templatePath) <> "" Then
        Set wordDoc = wordApp.documents.Open(templatePath)
        If Not wordDoc Is Nothing Then
            FillWordTemplate wordDoc
            GoTo SaveDocument
        End If
    End If
    
    Set wordDoc = wordApp.documents.Add
    CreateLetterFromScratch wordDoc
    
SaveDocument:
    Dim fileName As String
    fileName = GenerateFileNameWithExecutor( _
        IIf(txtAddressee Is Nothing, "Письмо", txtAddressee.Value), _
        IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value), _
        IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value))
    
    wordDoc.SaveAs fileName
    Debug.Print "Файл сохранен: " & fileName
    
    On Error Resume Next
    ThisWorkbook.Save
    Debug.Print "Книга Excel сохранена"
    On Error GoTo ErrorHandler
    
    wordApp.Visible = True
    wordDoc.Activate
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при создании письма: " & Err.description, vbCritical
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Private Sub FillWordTemplate(wordDoc As Object)
    On Error GoTo TemplateError
    
    Dim addresseeText As String, addressText As String, numberText As String
    Dim dateText As String, executorText As String, phoneText As String
    Dim letterText As String
    
    addresseeText = IIf(txtAddressee Is Nothing, "", txtAddressee.Value)
    addressText = FormatRecipientAddress(CreateAddressArray())
    numberText = IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value)
    
    dateText = FormatLetterDate(IIf(txtLetterDate Is Nothing, "", txtLetterDate.Value))
    Debug.Print "Отформатированная дата: " & dateText
    
    executorText = IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value)
    phoneText = GetExecutorPhone(executorText)
    letterText = GetDocumentTypeText(IIf(cmbDocumentType Is Nothing, "", cmbDocumentType.Value))
    
    SafeReplaceInWord wordDoc, "НаименованиеПолучателя", addresseeText
    SafeReplaceInWord wordDoc, "АдресПолучателя", addressText
    SafeReplaceInWord wordDoc, "НомерИсходящего", numberText
    SafeReplaceInWord wordDoc, "ДатаИсходящего", dateText
    SafeReplaceInWord wordDoc, "ИсполнительФИО", executorText
    SafeReplaceInWord wordDoc, "ТелефонИсполнителя", phoneText
    SafeReplaceInWord wordDoc, "ТекстПисьма", letterText
    
    ReplaceAttachmentsInTemplateWithFontAndSum wordDoc, 10
    
    Exit Sub
    
TemplateError:
    MsgBox "Ошибка заполнения шаблона: " & Err.description, vbCritical
End Sub

Private Sub CreateLetterFromScratch(wordDoc As Object)
    On Error GoTo ScratchError
    
    Dim content As String
    Dim letterText As String
    Dim addresseeText As String, addressText As String
    Dim numberText As String, dateText As String, executorText As String
    
    addresseeText = IIf(txtAddressee Is Nothing, "", txtAddressee.Value)
    addressText = FormatRecipientAddress(CreateAddressArray())
    numberText = IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value)
    dateText = IIf(txtLetterDate Is Nothing, "", txtLetterDate.Value)
    executorText = IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value)
    letterText = GetDocumentTypeText(IIf(cmbDocumentType Is Nothing, "", cmbDocumentType.Value))
    
    content = "Командиру войсковой части " & addresseeText & vbCrLf & vbCrLf
    content = content & addressText & vbCrLf & vbCrLf & vbCrLf
    content = content & letterText & vbCrLf & vbCrLf
    content = content & "Исполнитель: " & executorText & vbCrLf
    content = content & "Телефон: " & GetExecutorPhone(executorText) & vbCrLf
    content = content & "Исх. №: " & numberText & vbCrLf
    content = content & "Дата: " & dateText & vbCrLf & vbCrLf
    
    wordDoc.content.Text = content
    
    AppendAttachmentsToDocumentWithFontAndSum wordDoc, 10
    
    Exit Sub
    
ScratchError:
    MsgBox "Ошибка создания письма: " & Err.description, vbCritical
End Sub

'=====================================================================
'  НОВЫЕ ФУНКЦИИ ДЛЯ РАБОТЫ С ПРИЛОЖЕНИЯМИ И СУММОЙ В WORD
'=====================================================================
Private Sub ReplaceAttachmentsInTemplateWithFontAndSum(wordDoc As Object, fontSize As Integer)
    On Error Resume Next
    
    Dim rng As Object
    Set rng = wordDoc.content
    
    With rng.Find
        .ClearFormatting
        .Forward = True
        .Wrap = 1
        .Text = "СписокПриложений"
        
        If .Execute Then
            Dim startPos As Long
            startPos = rng.Start
            
            rng.Delete
            
            Dim attachmentFragments As Collection
            Set attachmentFragments = FormatAttachmentsListForWordWithSum(documentsList)
            
            Dim i As Long
            For i = 1 To attachmentFragments.count
                If i > 1 Then rng.InsertAfter vbCrLf
                rng.InsertAfter CStr(attachmentFragments(i))
                rng.Collapse 0
            Next i
            
            Dim attachmentRange As Object
            Set attachmentRange = wordDoc.Range(startPos, rng.End)
            
            FormatAttachmentsInWord attachmentRange, fontSize
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub AppendAttachmentsToDocumentWithFontAndSum(wordDoc As Object, fontSize As Integer)
    On Error Resume Next
    
    Dim rng As Object
    Set rng = wordDoc.content
    rng.Collapse 0
    
    rng.InsertAfter "Приложение: "
    
    Dim attachmentFragments As Collection
    Set attachmentFragments = FormatAttachmentsListForWordWithSum(documentsList)
    
    Dim startPos As Long
    startPos = rng.End
    
    Dim i As Long
    For i = 1 To attachmentFragments.count
        If i > 1 Then rng.InsertAfter vbCrLf
        rng.InsertAfter CStr(attachmentFragments(i))
        rng.Collapse 0
    Next i
    
    Dim attachmentRange As Object
    Set attachmentRange = wordDoc.Range(startPos, rng.End)
    
    FormatAttachmentsInWord attachmentRange, fontSize
    
    rng.InsertAfter vbCrLf & vbCrLf
    
    On Error GoTo 0
End Sub

Public Function FormatAttachmentsListForWordWithSum(documentsList As Collection) As Collection
    Set FormatAttachmentsListForWordWithSum = New Collection
    
    If documentsList Is Nothing Or documentsList.count = 0 Then
        FormatAttachmentsListForWordWithSum.Add "документы не указаны;"
        Exit Function
    End If
    
    Dim currentFragment As String
    Dim i As Long
    Dim docText As String
    
    For i = 1 To documentsList.count
        docText = i & "). " & FormatDocumentNameWithSum(documentsList(i)) & ";"
        
        If Len(currentFragment & vbCrLf & docText) > 180 Then
            If Len(currentFragment) > 0 Then
                FormatAttachmentsListForWordWithSum.Add currentFragment
                currentFragment = ""
            End If
        End If
        
        If Len(currentFragment) > 0 Then
            currentFragment = currentFragment & vbCrLf
        End If
        
        currentFragment = currentFragment & docText
    Next i
    
    If Len(currentFragment) > 0 Then
        FormatAttachmentsListForWordWithSum.Add currentFragment
    End If
End Function

'=====================================================================
'      СОХРАНЕНИЕ В БАЗУ ДАННЫХ С СУММОЙ
'=====================================================================
Private Sub SaveLetterToDatabase()
    Dim letterDate As Date
    
    On Error Resume Next
    If txtLetterDate Is Nothing Then
        letterDate = Date
    ElseIf IsDate(txtLetterDate.Value) Then
        letterDate = CDate(txtLetterDate.Value)
    Else
        letterDate = Date
    End If
    On Error GoTo 0
    
    SaveLetterInfoWithSum IIf(txtAddressee Is Nothing, "", txtAddressee.Value), _
                          IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value), _
                          letterDate, documentsList, _
                          IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value), _
                          IIf(cmbDocumentType Is Nothing, "", _
                              IIf(cmbDocumentType.ListIndex >= 0, cmbDocumentType.Value, ""))
End Sub

'=====================================================================
'      АВТООБНОВЛЕНИЕ АДРЕСОВ
'=====================================================================
Private Sub AutoUpdateAddressIfChanged()
    On Error Resume Next
    
    If selectedAddressRow <= 1 Then Exit Sub
    
    Dim currentAddress As Variant
    currentAddress = CreateAddressArray()
    
    If AddressDataHasChanges(selectedAddressRow, currentAddress) Then
        UpdateExistingAddress selectedAddressRow, currentAddress
        Debug.Print "Адрес автоматически обновлен в строке " & selectedAddressRow
        ClearAddressCache
    End If
    
    On Error GoTo 0
End Sub

Private Function AddressDataHasChanges(rowNumber As Long, newAddressArray As Variant) As Boolean
    AddressDataHasChanges = False
    
    On Error GoTo CompareError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Адреса")
    
    Dim i As Long, matchCount As Integer
    For i = 0 To UBound(newAddressArray)
        Dim sheetValue As String, formValue As String
        sheetValue = UCase(Trim(CStr(ws.Cells(rowNumber, i + 1).Value)))
        formValue = UCase(Trim(CStr(newAddressArray(i))))
        
        If sheetValue <> formValue Then
            Debug.Print "Изменение в столбце " & (i + 1) & ": '" & ws.Cells(rowNumber, i + 1).Value & "' -> '" & newAddressArray(i) & "'"
            AddressDataHasChanges = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
CompareError:
    Debug.Print "Ошибка сравнения данных адреса: " & Err.description
    AddressDataHasChanges = False
End Function

Private Sub ValidateAndUpdateSelectedAddress()
    On Error Resume Next
    
    If selectedAddressRow > 1 Then
        Dim city As String, region As String, postal As String
        city = Trim(Me.Controls("txtCity").Value)
        region = Trim(Me.Controls("txtRegion").Value)
        postal = Trim(Me.Controls("txtPostalCode").Value)
        
        If Len(city) > 0 And Len(region) > 0 And Len(postal) > 0 Then
            AutoUpdateAddressIfChanged
        End If
    End If
    
    On Error GoTo 0
End Sub

'=====================================================================
'      НАСТРОЙКА МНОГОСТРОЧНЫХ ПОЛЕЙ
'=====================================================================
Private Sub ConfigureMultilineTextBoxes()
    On Error Resume Next
    
    Dim ctrl As Control
    Dim textboxNames As Variant
    Dim i As Long
    
    textboxNames = Array("txtAddressee", "txtStreet", "txtCity", "txtDistrict", "txtRegion", "txtPostalCode")
    
    For i = LBound(textboxNames) To UBound(textboxNames)
        Set ctrl = Me.Controls(textboxNames(i))
        If Not ctrl Is Nothing Then
            ctrl.Multiline = True
            ctrl.WordWrap = True
            ctrl.ScrollBars = 2
            
            If ctrl.Height < 40 Then ctrl.Height = 35
        End If
    Next i
    
    On Error GoTo 0
End Sub

Private Sub AutoResizeTextBoxHeight(controlName As String)
    On Error Resume Next
    
    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    
    If Not ctrl Is Nothing Then
        Dim textLength As Long
        Dim linesCount As Long
        
        textLength = Len(ctrl.Value)
        linesCount = Int(textLength / 40) + 1
        
        If linesCount < 1 Then linesCount = 1
        If linesCount > 4 Then linesCount = 4
        
        ctrl.Height = linesCount * 18 + 10
    End If
    
    On Error GoTo 0
End Sub

'=====================================================================
'      БЕЗОПАСНАЯ УСТАНОВКА ФОКУСА
'=====================================================================
Private Sub SafeSetFocus(controlName As String)
    On Error Resume Next
    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then
        If ctrl.Enabled And ctrl.Visible Then
            ctrl.SetFocus
        End If
    End If
    On Error GoTo 0
End Sub

'=====================================================================
'      ЛОКАЛЬНЫЕ ФУНКЦИИ
'=====================================================================
Public Function TryParseDate(rawText As String, ByRef outDate As Date) As Boolean
    Dim t As String, d As Date, ok As Boolean
    TryParseDate = False
    
    If Len(Trim(rawText)) = 0 Then Exit Function
    
    If IsDate(rawText) Then
        outDate = CDate(rawText)
        TryParseDate = True
        Exit Function
    End If
    
    t = Replace(rawText, "/", ".")
    
    Dim clean As String, i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If IsNumeric(ch) Then clean = clean & ch
    Next i
    
    Select Case Len(clean)
        Case 8
            ok = IsDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & "." & Right(clean, 4))
            If ok Then outDate = CDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & "." & Right(clean, 4))
        Case 6
            ok = IsDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & ".20" & Right(clean, 2))
            If ok Then outDate = CDate(Left(clean, 2) & "." & Mid(clean, 3, 2) & ".20" & Right(clean, 2))
        Case 5
            ok = IsDate(Left(clean, 1) & "." & Mid(clean, 2, 2) & ".20" & Right(clean, 2))
            If ok Then outDate = CDate(Left(clean, 1) & "." & Mid(clean, 2, 2) & ".20" & Right(clean, 2))
        Case 4
            ok = IsDate(Left(clean, 2) & "." & Right(clean, 2) & "." & Year(Date))
            If ok Then outDate = CDate(Left(clean, 2) & "." & Right(clean, 2) & "." & Year(Date))
        Case Else
            ok = False
    End Select
    
    TryParseDate = ok
End Function

'=====================================================================
'  КНОПКИ ОТМЕНА И ЗАКРЫТИЕ
'=====================================================================
Private Sub btnCancel_Click()
    If MsgBox(T("dialog.cancel_letter_creation", "Отменить создание письма?"), vbYesNo + vbQuestion) = vbYes Then
        ClearCache
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If documentsList.count > 0 Then
        If MsgBox(T("dialog.discard_unsaved_documents", "Несохраненные документы будут потеряны. Закрыть?"), vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
        Else
            ClearCache
        End If
    Else
        ClearCache
    End If
End Sub


