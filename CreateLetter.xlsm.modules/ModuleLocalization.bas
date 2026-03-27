Attribute VB_Name = "ModuleLocalization"
' ======================================================================
' Module: ModuleLocalization
' Purpose: Provide workbook-backed localization helpers and built-in defaults for UI/runtime messages
' Version: 1.1.0 - 27.03.2026
' Notes:
' - This module does not change workbook schema.
' - Workbook sheet-based localization is optional and loaded only if present.
' - Built-in translations provide a safe fallback when the Localization sheet is incomplete.
' ======================================================================

Option Explicit

Private Const DEFAULT_LANGUAGE_CODE As String = "ru"
Private Const LOCALIZATION_SHEET_NAME As String = "Localization"

Private localizationCache As Object
Private activeLanguageCode As String

Public Function T(ByVal key As String, Optional ByVal fallback As String = "") As String
    T = Translate(key, fallback)
End Function

Public Function Translate(ByVal key As String, Optional ByVal fallback As String = "") As String
    Dim normalizedKey As String
    normalizedKey = NormalizeLocalizationKey(key)
    
    If Len(normalizedKey) = 0 Then
        Translate = fallback
        Exit Function
    End If
    
    EnsureLocalizationLoaded
    
    Dim resolvedText As String
    resolvedText = ResolveLocalizedText(GetAppLanguage(), normalizedKey)
    
    If Len(resolvedText) = 0 Then
        resolvedText = ResolveLocalizedText(DEFAULT_LANGUAGE_CODE, normalizedKey)
    End If
    
    If Len(resolvedText) = 0 Then
        If Len(fallback) > 0 Then
            Translate = fallback
        Else
            Translate = key
        End If
    Else
        Translate = resolvedText
    End If
End Function

Public Function TryGetLocalizedText(ByVal key As String, ByRef outText As String, Optional ByVal languageCode As String = "") As Boolean
    Dim requestedLanguage As String
    requestedLanguage = NormalizeLanguageCode(languageCode)
    If Len(requestedLanguage) = 0 Then
        requestedLanguage = GetAppLanguage()
    End If
    
    EnsureLocalizationLoaded
    
    outText = ResolveLocalizedText(requestedLanguage, NormalizeLocalizationKey(key))
    TryGetLocalizedText = (Len(outText) > 0)
End Function

Public Function GetAppLanguage() As String
    If Len(activeLanguageCode) = 0 Then
        activeLanguageCode = DEFAULT_LANGUAGE_CODE
    End If
    
    GetAppLanguage = activeLanguageCode
End Function

Public Sub SetAppLanguage(ByVal languageCode As String)
    Dim normalizedLanguage As String
    normalizedLanguage = NormalizeLanguageCode(languageCode)
    
    If Len(normalizedLanguage) = 0 Then
        normalizedLanguage = DEFAULT_LANGUAGE_CODE
    End If
    
    activeLanguageCode = normalizedLanguage
End Sub

Public Sub ResetLocalizationCache()
    If Not localizationCache Is Nothing Then
        localizationCache.RemoveAll
        Set localizationCache = Nothing
    End If
End Sub

Public Function LocalizationKeyExists(ByVal key As String) As Boolean
    EnsureLocalizationLoaded
    LocalizationKeyExists = localizationCache.Exists(BuildLocalizationMapKey(GetAppLanguage(), NormalizeLocalizationKey(key)))
End Function

Public Function GetLocalizationStats() As String
    EnsureLocalizationLoaded
    
    Dim stats As String
    stats = "Localization entries: " & localizationCache.Count & vbCrLf & _
            "Active language: " & GetAppLanguage() & vbCrLf & _
            "Default language: " & DEFAULT_LANGUAGE_CODE
    
    GetLocalizationStats = stats
End Function

Private Sub EnsureLocalizationLoaded()
    If localizationCache Is Nothing Then
        Set localizationCache = CreateObject("Scripting.Dictionary")
        localizationCache.CompareMode = vbTextCompare
        LoadBuiltInLocalization
        LoadLocalizationFromWorkbook
    End If
    
    If Len(activeLanguageCode) = 0 Then
        activeLanguageCode = DEFAULT_LANGUAGE_CODE
    End If
End Sub

Private Sub LoadBuiltInLocalization()
    ' Built-in defaults keep runtime localization safe even before workbook data is refreshed.
    AddTranslation "ru", "app.language.name", "Русский"
    AddTranslation "en", "app.language.name", "English"
    AddTranslation "ru", "common.ok", "ОК"
    AddTranslation "en", "common.ok", "OK"
    AddTranslation "ru", "common.cancel", "Отмена"
    AddTranslation "en", "common.cancel", "Cancel"
    AddTranslation "ru", "common.yes", "Да"
    AddTranslation "en", "common.yes", "Yes"
    AddTranslation "ru", "common.no", "Нет"
    AddTranslation "en", "common.no", "No"
    AddTranslation "ru", "common.of", "из"
    AddTranslation "en", "common.of", "of"
    AddTranslation "ru", "dialog.cancel_letter_creation", "Отменить создание письма?"
    AddTranslation "en", "dialog.cancel_letter_creation", "Cancel letter creation?"
    AddTranslation "ru", "dialog.discard_unsaved_documents", "Несохраненные документы будут потеряны. Закрыть?"
    AddTranslation "en", "dialog.discard_unsaved_documents", "Unsaved documents will be lost. Close?"
    AddTranslation "ru", "form.letter_creator.title", "Формирование писем"
    AddTranslation "en", "form.letter_creator.title", "Letter Builder"
    AddTranslation "ru", "form.letter_history.title", "История отправленных писем"
    AddTranslation "en", "form.letter_history.title", "Letter History"
    AddTranslation "ru", "form.letter_creator.tip.document_sum", "Сумма документа в рублях (необязательно). Например: 125000"
    AddTranslation "en", "form.letter_creator.tip.document_sum", "Document sum in rubles (optional). For example: 125000"
    AddTranslation "ru", "form.letter_creator.progress.page", "Шаг"
    AddTranslation "en", "form.letter_creator.progress.page", "Step"
    AddTranslation "ru", "form.letter_creator.attachments_count", "Выбрано документов:"
    AddTranslation "en", "form.letter_creator.attachments_count", "Selected documents:"
    AddTranslation "ru", "form.letter_creator.caption.edit_address", "Изменить адрес"
    AddTranslation "en", "form.letter_creator.caption.edit_address", "Edit address"
    AddTranslation "ru", "form.letter_creator.tip.edit_address", "Редактировать выбранный адрес"
    AddTranslation "en", "form.letter_creator.tip.edit_address", "Edit selected address"
    AddTranslation "ru", "form.letter_creator.caption.delete_address", "Удалить адрес"
    AddTranslation "en", "form.letter_creator.caption.delete_address", "Delete address"
    AddTranslation "ru", "form.letter_creator.tip.delete_address", "Удалить выбранный адрес"
    AddTranslation "en", "form.letter_creator.tip.delete_address", "Delete selected address"
    AddTranslation "ru", "form.letter_creator.tip.phone", "Телефон адресата (формат: 8-xxx-xxx-xx-xx)"
    AddTranslation "en", "form.letter_creator.tip.phone", "Addressee phone (format: 8-xxx-xxx-xx-xx)"
    AddTranslation "ru", "form.letter_creator.caption.letter_history", "История писем"
    AddTranslation "en", "form.letter_creator.caption.letter_history", "Letters History"
    AddTranslation "ru", "form.letter_creator.tip.letter_history", "Открыть форму истории отправленных писем"
    AddTranslation "en", "form.letter_creator.tip.letter_history", "Open sent letters history form"
    AddTranslation "ru", "form.letter_creator.tip.selected_attachments", "Для просмотра полного названия наведите на элемент"
    AddTranslation "en", "form.letter_creator.tip.selected_attachments", "Hover over the item to see the full name"
    AddTranslation "ru", "form.letter_creator.tip.address_search", "Введите часть наименования для поиска адресата"
    AddTranslation "en", "form.letter_creator.tip.address_search", "Enter part of the name to search for the addressee"
    AddTranslation "ru", "form.letter_creator.tip.letter_number", "Введите номер после 7/ (например: 125 > получится 7/125)"
    AddTranslation "en", "form.letter_creator.tip.letter_number", "Enter the number after 7/ (for example: 125 becomes 7/125)"
    AddTranslation "ru", "form.letter_creator.tip.letter_date", "Формат: дд.мм.гггг"
    AddTranslation "en", "form.letter_creator.tip.letter_date", "Format: dd.mm.yyyy"
    AddTranslation "ru", "form.letter_creator.caption.next", "Далее >"
    AddTranslation "en", "form.letter_creator.caption.next", "Next >"
    AddTranslation "ru", "form.letter_creator.caption.create_letter", "СОЗДАТЬ ПИСЬМО"
    AddTranslation "en", "form.letter_creator.caption.create_letter", "CREATE LETTER"
    AddTranslation "ru", "form.letter_creator.caption.back", "< Назад"
    AddTranslation "en", "form.letter_creator.caption.back", "< Back"
    AddTranslation "ru", "form.letter_creator.caption.cancel", "Отмена"
    AddTranslation "en", "form.letter_creator.caption.cancel", "Cancel"
    AddTranslation "ru", "form.letter_creator.caption.save_address", "Сохранить адрес"
    AddTranslation "en", "form.letter_creator.caption.save_address", "Save address"
    AddTranslation "ru", "form.letter_creator.caption.clear_search", "Очистить"
    AddTranslation "en", "form.letter_creator.caption.clear_search", "Clear"
    AddTranslation "ru", "form.letter_creator.page.step_1", "Шаг 1: Адресат"
    AddTranslation "en", "form.letter_creator.page.step_1", "Step 1: Addressee"
    AddTranslation "ru", "form.letter_creator.page.step_2", "Шаг 2: Письмо"
    AddTranslation "en", "form.letter_creator.page.step_2", "Step 2: Letter"
    AddTranslation "ru", "form.letter_creator.page.step_3", "Шаг 3: Приложения"
    AddTranslation "en", "form.letter_creator.page.step_3", "Step 3: Attachments"
    AddTranslation "ru", "form.letter_creator.page.step_4", "Шаг 4: Создание"
    AddTranslation "en", "form.letter_creator.page.step_4", "Step 4: Create"
    AddTranslation "ru", "form.letter_creator.label.stage", "Этап:"
    AddTranslation "en", "form.letter_creator.label.stage", "Stage:"
    AddTranslation "ru", "form.letter_creator.label.current_action", "Текущее действие"
    AddTranslation "en", "form.letter_creator.label.current_action", "Current action"
    AddTranslation "ru", "form.letter_creator.label.search_addressee", "Поиск существующего адресата"
    AddTranslation "en", "form.letter_creator.label.search_addressee", "Search existing addressee"
    AddTranslation "ru", "form.letter_creator.label.city", "Город"
    AddTranslation "en", "form.letter_creator.label.city", "City"
    AddTranslation "ru", "form.letter_creator.label.district", "Район"
    AddTranslation "en", "form.letter_creator.label.district", "District"
    AddTranslation "ru", "form.letter_creator.label.region", "Регион"
    AddTranslation "en", "form.letter_creator.label.region", "Region"
    AddTranslation "ru", "form.letter_creator.label.postal_code", "Почтовый индекс"
    AddTranslation "en", "form.letter_creator.label.postal_code", "Postal code"
    AddTranslation "ru", "form.letter_creator.label.executor", "Исполнитель"
    AddTranslation "en", "form.letter_creator.label.executor", "Executor"
    AddTranslation "ru", "form.letter_creator.label.letter_date", "Дата письма"
    AddTranslation "en", "form.letter_creator.label.letter_date", "Letter date"
    AddTranslation "ru", "form.letter_creator.label.letter_number", "Номер письма"
    AddTranslation "en", "form.letter_creator.label.letter_number", "Letter number"
    AddTranslation "ru", "form.letter_creator.label.search_attachment", "Поиск приложения"
    AddTranslation "en", "form.letter_creator.label.search_attachment", "Search attachment"
    AddTranslation "ru", "form.letter_creator.label.selected_attachments", "Выбранные приложения"
    AddTranslation "en", "form.letter_creator.label.selected_attachments", "Selected attachments"
    AddTranslation "ru", "form.letter_creator.label.document_ownership", "Принадлежность документа"
    AddTranslation "en", "form.letter_creator.label.document_ownership", "Document ownership"
    AddTranslation "ru", "form.letter_creator.label.date", "Дата"
    AddTranslation "en", "form.letter_creator.label.date", "Date"
    AddTranslation "ru", "form.letter_creator.label.copies", "Экземпляры"
    AddTranslation "en", "form.letter_creator.label.copies", "Copies"
    AddTranslation "ru", "form.letter_creator.label.sheets", "Листы"
    AddTranslation "en", "form.letter_creator.label.sheets", "Sheets"
    AddTranslation "ru", "form.letter_creator.label.found_addresses", "Найденные адреса"
    AddTranslation "en", "form.letter_creator.label.found_addresses", "Found addresses"
    AddTranslation "ru", "form.letter_creator.label.street_house", "Улица, дом"
    AddTranslation "en", "form.letter_creator.label.street_house", "Street, house"
    AddTranslation "ru", "form.letter_creator.label.addressee", "Адресат"
    AddTranslation "en", "form.letter_creator.label.addressee", "Addressee"
    AddTranslation "ru", "form.letter_creator.label.available_attachments", "Доступные приложения"
    AddTranslation "en", "form.letter_creator.label.available_attachments", "Available attachments"
    AddTranslation "ru", "form.letter_creator.label.number", "Номер"
    AddTranslation "en", "form.letter_creator.label.number", "Number"
    AddTranslation "ru", "form.letter_creator.label.summary_addressee", "Адресат:"
    AddTranslation "en", "form.letter_creator.label.summary_addressee", "Addressee:"
    AddTranslation "ru", "form.letter_creator.label.summary_letter_number", "Номер письма:"
    AddTranslation "en", "form.letter_creator.label.summary_letter_number", "Letter number:"
    AddTranslation "ru", "form.letter_creator.label.summary_date", "Дата:"
    AddTranslation "en", "form.letter_creator.label.summary_date", "Date:"
    AddTranslation "ru", "form.letter_creator.label.summary_executor", "Исполнитель:"
    AddTranslation "en", "form.letter_creator.label.summary_executor", "Executor:"
    AddTranslation "ru", "form.letter_creator.label.summary_document_count", "Количество документов:"
    AddTranslation "en", "form.letter_creator.label.summary_document_count", "Document count:"
    AddTranslation "ru", "form.letter_creator.label.summary_attachments", "Приложения:"
    AddTranslation "en", "form.letter_creator.label.summary_attachments", "Attachments:"
    AddTranslation "ru", "form.letter_creator.label.document_sum", "Сумма документа"
    AddTranslation "en", "form.letter_creator.label.document_sum", "Document sum"
    AddTranslation "ru", "form.letter_creator.label.selected_document", "Выбранный документ:"
    AddTranslation "en", "form.letter_creator.label.selected_document", "Selected document:"
    AddTranslation "ru", "form.letter_creator.frame.address_details", "Данные адреса"
    AddTranslation "en", "form.letter_creator.frame.address_details", "Address details"
    AddTranslation "ru", "form.letter_creator.frame.letter_summary", "Сводка письма"
    AddTranslation "en", "form.letter_creator.frame.letter_summary", "Letter summary"
    AddTranslation "ru", "form.letter_creator.msg.history_open_error", "Ошибка при открытии формы истории: "
    AddTranslation "en", "form.letter_creator.msg.history_open_error", "Error opening history form: "
    AddTranslation "ru", "form.letter_creator.msg.letter_created", "Письмо успешно создано!"
    AddTranslation "en", "form.letter_creator.msg.letter_created", "Letter created successfully!"
    AddTranslation "ru", "form.letter_creator.msg.address_saved", "Адрес сохранен."
    AddTranslation "en", "form.letter_creator.msg.address_saved", "Address saved."
    AddTranslation "ru", "form.letter_creator.msg.address_updated", "Адрес успешно обновлен."
    AddTranslation "en", "form.letter_creator.msg.address_updated", "Address updated successfully."
    AddTranslation "ru", "form.letter_creator.msg.address_delete_confirm", "Вы уверены, что хотите удалить этот адрес?"
    AddTranslation "en", "form.letter_creator.msg.address_delete_confirm", "Are you sure you want to delete this address?"
    AddTranslation "ru", "form.letter_creator.msg.address_deleted", "Адрес успешно удален."
    AddTranslation "en", "form.letter_creator.msg.address_deleted", "Address deleted successfully."
    AddTranslation "ru", "form.letter_creator.msg.address_delete_error", "Ошибка при удалении адреса: "
    AddTranslation "en", "form.letter_creator.msg.address_delete_error", "Error deleting address: "
    AddTranslation "ru", "form.letter_creator.msg.select_document_left", "Выберите документ в левом списке."
    AddTranslation "en", "form.letter_creator.msg.select_document_left", "Select a document in the left list."
    AddTranslation "ru", "form.letter_creator.msg.select_document_right", "Выберите документ в правом списке."
    AddTranslation "en", "form.letter_creator.msg.select_document_right", "Select a document in the right list."
    AddTranslation "ru", "form.letter_creator.msg.duplicate_document_error", "Ошибка при дублировании документа: "
    AddTranslation "en", "form.letter_creator.msg.duplicate_document_error", "Error duplicating document: "
    AddTranslation "ru", "form.letter_creator.msg.create_letter_error", "Ошибка при создании письма: "
    AddTranslation "en", "form.letter_creator.msg.create_letter_error", "Error creating letter: "
    AddTranslation "ru", "form.letter_creator.menu.document_actions_title", "Действия с документом"
    AddTranslation "en", "form.letter_creator.menu.document_actions_title", "Document actions"
    AddTranslation "ru", "form.letter_creator.menu.document_actions_prompt", "Выберите действие:"
    AddTranslation "en", "form.letter_creator.menu.document_actions_prompt", "Select action:"
    AddTranslation "ru", "form.letter_creator.menu.document_action.edit", "1 - Изменить реквизиты"
    AddTranslation "en", "form.letter_creator.menu.document_action.edit", "1 - Edit details"
    AddTranslation "ru", "form.letter_creator.menu.document_action.duplicate", "2 - Дублировать документ"
    AddTranslation "en", "form.letter_creator.menu.document_action.duplicate", "2 - Duplicate document"
    AddTranslation "ru", "form.letter_creator.menu.document_action.remove", "3 - Удалить из списка"
    AddTranslation "en", "form.letter_creator.menu.document_action.remove", "3 - Remove from list"
    AddTranslation "ru", "form.letter_creator.menu.document_action.move_up", "4 - Переместить вверх"
    AddTranslation "en", "form.letter_creator.menu.document_action.move_up", "4 - Move up"
    AddTranslation "ru", "form.letter_creator.menu.document_action.move_down", "5 - Переместить вниз"
    AddTranslation "en", "form.letter_creator.menu.document_action.move_down", "5 - Move down"
    AddTranslation "ru", "validation.creator.page.addressee_required", "Заполните поле 'Адресат'."
    AddTranslation "en", "validation.creator.page.addressee_required", "Fill in the 'Addressee' field."
    AddTranslation "ru", "validation.creator.page.city_required", "Заполните поле 'Город'. Это обязательное поле."
    AddTranslation "en", "validation.creator.page.city_required", "Fill in the 'City' field. This field is required."
    AddTranslation "ru", "validation.creator.page.region_required", "Заполните поле 'Регион'. Это обязательное поле."
    AddTranslation "en", "validation.creator.page.region_required", "Fill in the 'Region' field. This field is required."
    AddTranslation "ru", "validation.creator.page.postal_code_required", "Заполните поле 'Почтовый индекс'. Это обязательное поле."
    AddTranslation "en", "validation.creator.page.postal_code_required", "Fill in the 'Postal code' field. This field is required."
    AddTranslation "ru", "validation.creator.page.phone_invalid", "Введите корректный номер телефона адресата."
    AddTranslation "en", "validation.creator.page.phone_invalid", "Enter a valid addressee phone number."
    AddTranslation "ru", "validation.creator.page.letter_number_required", "Введите номер исходящего письма."
    AddTranslation "en", "validation.creator.page.letter_number_required", "Enter the outgoing letter number."
    AddTranslation "ru", "validation.creator.page.letter_date_required", "Введите дату письма."
    AddTranslation "en", "validation.creator.page.letter_date_required", "Enter the letter date."
    AddTranslation "ru", "validation.creator.page.executor_required", "Выберите исполнителя. Это обязательное поле."
    AddTranslation "en", "validation.creator.page.executor_required", "Select an executor. This field is required."
    AddTranslation "ru", "validation.creator.page.letter_date_invalid", "Неверный формат даты письма."
    AddTranslation "en", "validation.creator.page.letter_date_invalid", "Invalid letter date format."
    AddTranslation "ru", "validation.creator.page.document_required", "Добавьте хотя бы один документ-приложение."
    AddTranslation "en", "validation.creator.page.document_required", "Add at least one attachment document."
    AddTranslation "ru", "validation.creator.submit.addressee_required", "Адресат не заполнен."
    AddTranslation "en", "validation.creator.submit.addressee_required", "Addressee is not filled in."
    AddTranslation "ru", "validation.creator.submit.city_required", "Город не заполнен."
    AddTranslation "en", "validation.creator.submit.city_required", "City is not filled in."
    AddTranslation "ru", "validation.creator.submit.region_required", "Регион не заполнен."
    AddTranslation "en", "validation.creator.submit.region_required", "Region is not filled in."
    AddTranslation "ru", "validation.creator.submit.postal_code_required", "Почтовый индекс не заполнен."
    AddTranslation "en", "validation.creator.submit.postal_code_required", "Postal code is not filled in."
    AddTranslation "ru", "validation.creator.submit.letter_number_required", "Номер письма не заполнен."
    AddTranslation "en", "validation.creator.submit.letter_number_required", "Letter number is not filled in."
    AddTranslation "ru", "validation.creator.submit.letter_date_required", "Дата письма не заполнена."
    AddTranslation "en", "validation.creator.submit.letter_date_required", "Letter date is not filled in."
    AddTranslation "ru", "validation.creator.submit.executor_required", "Исполнитель не выбран."
    AddTranslation "en", "validation.creator.submit.executor_required", "Executor is not selected."
    AddTranslation "ru", "validation.creator.submit.document_required", "Добавьте хотя бы один документ."
    AddTranslation "en", "validation.creator.submit.document_required", "Add at least one document."
    AddTranslation "ru", "validation.address.record_invalid", "Неверный формат записи адреса."
    AddTranslation "en", "validation.address.record_invalid", "Invalid address record format."
    AddTranslation "ru", "validation.address.record_incomplete", "Данные адреса неполные."
    AddTranslation "en", "validation.address.record_incomplete", "Address data is incomplete."
    AddTranslation "ru", "validation.address.row_invalid", "Ссылка на строку адреса недопустима."
    AddTranslation "en", "validation.address.row_invalid", "Address row reference is invalid."
    AddTranslation "ru", "validation.address.create.addressee_required", "Введите имя адресата."
    AddTranslation "en", "validation.address.create.addressee_required", "Enter the addressee name."
    AddTranslation "ru", "validation.address.create.duplicate", "Такой адрес уже существует."
    AddTranslation "en", "validation.address.create.duplicate", "This address already exists."
    AddTranslation "ru", "validation.address.edit.selection_required", "Выберите адрес для редактирования."
    AddTranslation "en", "validation.address.edit.selection_required", "Select an address to edit."
    AddTranslation "ru", "validation.address.edit.duplicate", "Адрес с такими данными уже существует."
    AddTranslation "en", "validation.address.edit.duplicate", "An address with the same data already exists."
    AddTranslation "ru", "validation.address.delete.selection_required", "Выберите адрес для удаления."
    AddTranslation "en", "validation.address.delete.selection_required", "Select an address to delete."
    AddTranslation "ru", "status.ready", "Готово"
    AddTranslation "en", "status.ready", "Ready"
    AddTranslation "ru", "status.searching", "Поиск..."
    AddTranslation "en", "status.searching", "Searching..."
    AddTranslation "ru", "error.generic", "Произошла ошибка."
    AddTranslation "en", "error.generic", "An error occurred."
End Sub

Private Sub LoadLocalizationFromWorkbook()
    On Error GoTo LoadError
    
    Dim ws As Worksheet
    Set ws = Nothing
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOCALIZATION_SHEET_NAME)
    On Error GoTo LoadError
    
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Dim entryKey As String
        Dim entryRu As String
        Dim entryEn As String
        
        entryKey = NormalizeLocalizationKey(CStr(ws.Cells(rowIndex, 1).Value))
        If Len(entryKey) = 0 Then GoTo NextRow
        
        entryRu = CStr(ws.Cells(rowIndex, 2).Value)
        entryEn = CStr(ws.Cells(rowIndex, 3).Value)
        
        If Len(entryRu) > 0 Then AddTranslation "ru", entryKey, entryRu
        If Len(entryEn) > 0 Then AddTranslation "en", entryKey, entryEn
        
NextRow:
    Next rowIndex
    
    Exit Sub
    
LoadError:
    Debug.Print "Localization load skipped: " & Err.Description
End Sub

Private Sub AddTranslation(ByVal languageCode As String, ByVal key As String, ByVal value As String)
    Dim mapKey As String
    mapKey = BuildLocalizationMapKey(languageCode, key)
    
    If localizationCache.Exists(mapKey) Then
        localizationCache(mapKey) = value
    Else
        localizationCache.Add mapKey, value
    End If
End Sub

Private Function ResolveLocalizedText(ByVal languageCode As String, ByVal key As String) As String
    Dim mapKey As String
    mapKey = BuildLocalizationMapKey(languageCode, key)
    
    If localizationCache.Exists(mapKey) Then
        ResolveLocalizedText = CStr(localizationCache(mapKey))
    Else
        ResolveLocalizedText = ""
    End If
End Function

Private Function BuildLocalizationMapKey(ByVal languageCode As String, ByVal key As String) As String
    BuildLocalizationMapKey = NormalizeLanguageCode(languageCode) & "|" & NormalizeLocalizationKey(key)
End Function

Private Function NormalizeLanguageCode(ByVal languageCode As String) As String
    NormalizeLanguageCode = LCase$(Trim$(languageCode))
End Function

Private Function NormalizeLocalizationKey(ByVal key As String) As String
    NormalizeLocalizationKey = LCase$(Trim$(key))
End Function
