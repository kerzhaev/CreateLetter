Attribute VB_Name = "ModuleLocalization"
' ======================================================================
' Module: ModuleLocalization
' Purpose: Provide workbook-backed localization helpers and built-in defaults for UI/runtime messages
' Version: 1.4.5 - 29.03.2026
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

Public Function t(ByVal key As String, Optional ByVal fallback As String = "") As String
    t = Translate(key, fallback)
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
    stats = "Localization entries: " & localizationCache.count & vbCrLf & _
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
    AddTranslation "ru", "common.not_specified", "Не указано"
    AddTranslation "en", "common.not_specified", "Not specified"
    AddTranslation "ru", "common.unknown_user", "Неизвестный пользователь"
    AddTranslation "en", "common.unknown_user", "Unknown user"
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
    AddTranslation "ru", "form.letter_creator.option.document_type.confirmed", "Чужие подтвержденные документы"
    AddTranslation "en", "form.letter_creator.option.document_type.confirmed", "Third-party confirmed documents"
    AddTranslation "ru", "form.letter_creator.option.document_type.own_confirmation", "Свои для подтверждения"
    AddTranslation "en", "form.letter_creator.option.document_type.own_confirmation", "Own for confirmation"
    AddTranslation "ru", "form.letter_creator.option.letter_type.regular", "Обычное"
    AddTranslation "en", "form.letter_creator.option.letter_type.regular", "Regular"
    AddTranslation "ru", "form.letter_creator.option.letter_type.fou", "ДСП"
    AddTranslation "en", "form.letter_creator.option.letter_type.fou", "FOU (For Official Use)"
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
    AddTranslation "ru", "form.letter_history.frame.search", "Поиск"
    AddTranslation "en", "form.letter_history.frame.search", "Search"
    AddTranslation "ru", "form.letter_history.frame.letters_list", "Список писем"
    AddTranslation "en", "form.letter_history.frame.letters_list", "Letter list"
    AddTranslation "ru", "form.letter_history.frame.status_update", "Обновление статуса"
    AddTranslation "en", "form.letter_history.frame.status_update", "Status update"
    AddTranslation "ru", "form.letter_history.frame.actions", "Действия"
    AddTranslation "en", "form.letter_history.frame.actions", "Actions"
    AddTranslation "ru", "form.letter_history.label.search_letters", "Поиск писем по истории доставки"
    AddTranslation "en", "form.letter_history.label.search_letters", "Search letters by delivery history"
    AddTranslation "ru", "form.letter_history.label.search_status", "Статус поиска"
    AddTranslation "en", "form.letter_history.label.search_status", "Search status"
    AddTranslation "ru", "form.letter_history.label.return_date", "Дата возврата"
    AddTranslation "en", "form.letter_history.label.return_date", "Return date"
    AddTranslation "ru", "form.letter_history.label.amount", "Сумма"
    AddTranslation "en", "form.letter_history.label.amount", "Amount"
    AddTranslation "ru", "form.letter_history.caption.update_status", "Обновить статус"
    AddTranslation "en", "form.letter_history.caption.update_status", "Update status"
    AddTranslation "ru", "form.letter_history.caption.close", "Закрыть"
    AddTranslation "en", "form.letter_history.caption.close", "Close"
    AddTranslation "ru", "form.letter_history.caption.refresh_data", "Обновить данные"
    AddTranslation "en", "form.letter_history.caption.refresh_data", "Refresh data"
    AddTranslation "ru", "form.letter_history.caption.clear_search", "Очистить поиск"
    AddTranslation "en", "form.letter_history.caption.clear_search", "Clear search"
    AddTranslation "ru", "form.letter_history.caption.clearing", "Очистка..."
    AddTranslation "en", "form.letter_history.caption.clearing", "Clearing..."
    AddTranslation "ru", "form.letter_history.caption.export_to_excel", "Экспорт в Excel"
    AddTranslation "en", "form.letter_history.caption.export_to_excel", "Export to Excel"
    AddTranslation "ru", "form.letter_history.caption.go_to_record", "Перейти к записи"
    AddTranslation "en", "form.letter_history.caption.go_to_record", "Go to record"
    AddTranslation "ru", "form.letter_history.caption.search_help", "Справка по поиску"
    AddTranslation "en", "form.letter_history.caption.search_help", "Search help"
    AddTranslation "ru", "form.letter_history.caption.received_back", "Документ получен обратно"
    AddTranslation "en", "form.letter_history.caption.received_back", "Document received back"
    AddTranslation "ru", "form.letter_history.tip.return_date", "Дата возврата документа (дд.мм.гггг)"
    AddTranslation "en", "form.letter_history.tip.return_date", "Document return date (dd.mm.yyyy)"
    AddTranslation "ru", "form.letter_history.tip.document_sum", "Сумма документа в рублях (только цифры или краткий комментарий)"
    AddTranslation "en", "form.letter_history.tip.document_sum", "Document sum in rubles (numbers only or brief comment)"
    AddTranslation "ru", "form.letter_history.tip.search", "Поиск по адресату, номеру, дате, приложениям, исполнителю" & vbCrLf & "Для поиска по сумме вводите только цифры (например: 125000)"
    AddTranslation "en", "form.letter_history.tip.search", "Search by addressee, number, date, attachments, executor" & vbCrLf & "To search by sum, enter numbers only (e.g.: 125000)"
    AddTranslation "ru", "form.letter_history.tip.double_click", "Дважды щелкните по письму, чтобы перейти к записи в таблице"
    AddTranslation "en", "form.letter_history.tip.double_click", "Double click on a letter to jump to the table record"
    AddTranslation "ru", "form.letter_history.msg.no_data", "На листе 'Letters' данные не найдены"
    AddTranslation "en", "form.letter_history.msg.no_data", "No data found in worksheet 'Letters'"
    AddTranslation "ru", "form.letter_history.msg.letters_loaded", "Писем загружено: "
    AddTranslation "en", "form.letter_history.msg.letters_loaded", "Letters loaded: "
    AddTranslation "ru", "form.letter_history.msg.showing_all", "Показаны все письма: "
    AddTranslation "en", "form.letter_history.msg.showing_all", "Showing all letters: "
    AddTranslation "ru", "form.letter_history.msg.searching_amount", "Ищем номер "
    AddTranslation "en", "form.letter_history.msg.searching_amount", "Searching for number "
    AddTranslation "ru", "form.letter_history.msg.searching_amount_suffix", " в суммах документов..."
    AddTranslation "en", "form.letter_history.msg.searching_amount_suffix", " in document amounts..."
    AddTranslation "ru", "form.letter_history.msg.letters_found", "Найдено писем: "
    AddTranslation "en", "form.letter_history.msg.letters_found", "Letters found: "
    AddTranslation "ru", "form.letter_history.msg.out_of", " из "
    AddTranslation "en", "form.letter_history.msg.out_of", " of "
    AddTranslation "ru", "form.letter_history.msg.select_record", "Выберите письмо для перехода к записи."
    AddTranslation "en", "form.letter_history.msg.select_record", "Select a letter to navigate to the record."
    AddTranslation "ru", "form.letter_history.msg.navigation_error_title", "Ошибка перехода"
    AddTranslation "en", "form.letter_history.msg.navigation_error_title", "Navigation error"
    AddTranslation "ru", "form.letter_history.msg.letters_sheet_missing", "Лист 'Letters' не найден."
    AddTranslation "en", "form.letter_history.msg.letters_sheet_missing", "Worksheet 'Letters' not found."
    AddTranslation "ru", "form.letter_history.msg.navigation_error", "Ошибка перехода к записи: "
    AddTranslation "en", "form.letter_history.msg.navigation_error", "Error navigating to record: "
    AddTranslation "ru", "form.letter_history.msg.selected_record", "Выбрана запись: "
    AddTranslation "en", "form.letter_history.msg.selected_record", "Selected record: "
    AddTranslation "ru", "form.letter_history.msg.select_status_update", "Выберите письмо для обновления статуса."
    AddTranslation "en", "form.letter_history.msg.select_status_update", "Select a letter to update the status."
    AddTranslation "ru", "form.letter_history.msg.status_updated", "Статус письма успешно обновлен."
    AddTranslation "en", "form.letter_history.msg.status_updated", "Letter status updated successfully."
    AddTranslation "ru", "form.letter_history.msg.data_refreshed", "Данные обновлены."
    AddTranslation "en", "form.letter_history.msg.data_refreshed", "Data refreshed."
    AddTranslation "ru", "form.letter_history.msg.no_export_data", "Нет данных для экспорта."
    AddTranslation "en", "form.letter_history.msg.no_export_data", "No data to export."
    AddTranslation "ru", "form.letter_history.msg.export_completed", "Экспорт завершен."
    AddTranslation "en", "form.letter_history.msg.export_completed", "Export completed."
    AddTranslation "ru", "form.letter_history.msg.records_exported", "Экспортировано записей: "
    AddTranslation "en", "form.letter_history.msg.records_exported", "Records exported: "
    AddTranslation "ru", "form.letter_history.msg.export_title", "Экспорт данных"
    AddTranslation "en", "form.letter_history.msg.export_title", "Data export"
    AddTranslation "ru", "form.letter_history.msg.export_error", "Ошибка экспорта: "
    AddTranslation "en", "form.letter_history.msg.export_error", "Export error: "
    AddTranslation "ru", "form.letter_history.export.header.addressee", "Адресат"
    AddTranslation "en", "form.letter_history.export.header.addressee", "Addressee"
    AddTranslation "ru", "form.letter_history.export.header.outgoing_number", "Исходящий номер"
    AddTranslation "en", "form.letter_history.export.header.outgoing_number", "Outgoing Number"
    AddTranslation "ru", "form.letter_history.export.header.outgoing_date", "Исходящая дата"
    AddTranslation "en", "form.letter_history.export.header.outgoing_date", "Outgoing Date"
    AddTranslation "ru", "form.letter_history.export.header.attachment_name", "Наименование приложения"
    AddTranslation "en", "form.letter_history.export.header.attachment_name", "Attachment Name"
    AddTranslation "ru", "form.letter_history.export.header.document_sum", "Сумма документа"
    AddTranslation "en", "form.letter_history.export.header.document_sum", "Document Sum"
    AddTranslation "ru", "form.letter_history.export.header.return_mark", "Отметка возврата"
    AddTranslation "en", "form.letter_history.export.header.return_mark", "Return Mark"
    AddTranslation "ru", "form.letter_history.export.header.executor_name", "Имя исполнителя"
    AddTranslation "en", "form.letter_history.export.header.executor_name", "Executor Name"
    AddTranslation "ru", "form.letter_history.export.header.send_type", "Тип отправки"
    AddTranslation "en", "form.letter_history.export.header.send_type", "Send Type"
    AddTranslation "ru", "form.letter_history.export.sheet_name", "История писем "
    AddTranslation "en", "form.letter_history.export.sheet_name", "Letters history "
    AddTranslation "ru", "form.letter_history.msg.invalid_date", "Неверный формат даты. Используйте дд.мм.гггг."
    AddTranslation "en", "form.letter_history.msg.invalid_date", "Invalid date format. Use dd.mm.yyyy."
    AddTranslation "ru", "form.letter_history.msg.search_hints_title", "Подсказки по поиску"
    AddTranslation "en", "form.letter_history.msg.search_hints_title", "Search Help"
    AddTranslation "ru", "form.letter_history.msg.search_hints_body", "ПОДСКАЗКИ ПО ПОИСКУ:" & vbCrLf & vbCrLf & "• Для поиска по сумме вводите только цифры: 125000" & vbCrLf & "• Система найдет '125000', '125 000', '125000 руб.'" & vbCrLf & "• Поиск работает одновременно по всем столбцам" & vbCrLf & "• Можно искать по части слова или числа" & vbCrLf & vbCrLf & "Нажмите 'Обновить данные', если вы вручную изменяли Excel"
    AddTranslation "en", "form.letter_history.msg.search_hints_body", "SEARCH HINTS:" & vbCrLf & vbCrLf & "• To search for a sum, enter only numbers: 125000" & vbCrLf & "• The system will find '125000', '125 000', '125000 rub.'" & vbCrLf & "• Search works across all columns simultaneously" & vbCrLf & "• You can search by part of a word or number" & vbCrLf & vbCrLf & "Click 'Refresh data' if you modified Excel manually"
    AddTranslation "ru", "history.status.received_suffix", " получено"
    AddTranslation "en", "history.status.received_suffix", " received"
    AddTranslation "ru", "history.status.not_received", "не получено"
    AddTranslation "en", "history.status.not_received", "not received"
    AddTranslation "ru", "history.status.received_label", "Получено "
    AddTranslation "en", "history.status.received_label", "Received "
    AddTranslation "ru", "history.status.pending_label", "Ожидается "
    AddTranslation "en", "history.status.pending_label", "Pending "
    AddTranslation "ru", "form.letter_history.msg.already_open", "Форма истории писем уже открыта!"
    AddTranslation "en", "form.letter_history.msg.already_open", "Letter history form is already open!"
    AddTranslation "ru", "form.letter_history.msg.open_error", "Ошибка при открытии формы истории писем: "
    AddTranslation "en", "form.letter_history.msg.open_error", "Error opening letter history form: "
    AddTranslation "ru", "core.runtime.error.invalid_data_format", "Ошибка: неверный формат данных"
    AddTranslation "en", "core.runtime.error.invalid_data_format", "Error: invalid data format"
    AddTranslation "ru", "core.address.error.save", "Ошибка при сохранении адреса: "
    AddTranslation "en", "core.address.error.save", "Error saving address: "
    AddTranslation "ru", "core.address.error.update", "Ошибка при обновлении адреса: "
    AddTranslation "en", "core.address.error.update", "Error updating address: "
    AddTranslation "ru", "core.address.error.delete", "Ошибка при удалении адреса: "
    AddTranslation "en", "core.address.error.delete", "Error deleting address: "
    AddTranslation "ru", "core.letter.error.save_info", "Ошибка при сохранении информации о письме: "
    AddTranslation "en", "core.letter.error.save_info", "Error saving letter info: "
    AddTranslation "ru", "core.attachments.not_specified_word", "документы не указаны;"
    AddTranslation "en", "core.attachments.not_specified_word", "documents not specified;"
    AddTranslation "ru", "core.attachments.not_specified", "Документы не указаны"
    AddTranslation "en", "core.attachments.not_specified", "Documents not specified"
    AddTranslation "ru", "core.document.label.number_prefix", " №"
    AddTranslation "en", "core.document.label.number_prefix", " No."
    AddTranslation "ru", "core.document.label.date_prefix", " от "
    AddTranslation "en", "core.document.label.date_prefix", " dated "
    AddTranslation "ru", "core.document.label.amount_prefix", " на сумму "
    AddTranslation "en", "core.document.label.amount_prefix", " for the amount of "
    AddTranslation "ru", "core.document.label.amount_suffix", " руб."
    AddTranslation "en", "core.document.label.amount_suffix", " rub."
    AddTranslation "ru", "core.document.label.copies_suffix", " экз."
    AddTranslation "en", "core.document.label.copies_suffix", " copies"
    AddTranslation "ru", "core.document.label.sheets_suffix", " л."
    AddTranslation "en", "core.document.label.sheets_suffix", " sheets"
    AddTranslation "ru", "core.letter.default_file_name", "Письмо"
    AddTranslation "en", "core.letter.default_file_name", "Letter"
    AddTranslation "ru", "core.letter.warning.workbook_not_saved", "Файл письма создан, но книгу не удалось сохранить: "
    AddTranslation "en", "core.letter.warning.workbook_not_saved", "Letter file was created, but the workbook was not saved: "
    AddTranslation "ru", "core.letter.error.create_document", "Ошибка при создании письма: "
    AddTranslation "en", "core.letter.error.create_document", "Error creating letter: "
    AddTranslation "ru", "core.letter.error.template_fill", "Ошибка заполнения шаблона: "
    AddTranslation "en", "core.letter.error.template_fill", "Template filling error: "
    AddTranslation "ru", "core.letter.error.create_fallback", "Ошибка при создании текста письма: "
    AddTranslation "en", "core.letter.error.create_fallback", "Letter creation error: "
    AddTranslation "ru", "core.letter.fallback.to_commander", "Командиру войсковой части "
    AddTranslation "en", "core.letter.fallback.to_commander", "To the Commander of military unit "
    AddTranslation "ru", "core.letter.fallback.executor", "Исполнитель: "
    AddTranslation "en", "core.letter.fallback.executor", "Executor: "
    AddTranslation "ru", "core.letter.fallback.phone", "Телефон: "
    AddTranslation "en", "core.letter.fallback.phone", "Phone: "
    AddTranslation "ru", "core.letter.fallback.ref_no", "Исх. №: "
    AddTranslation "en", "core.letter.fallback.ref_no", "Ref. No.: "
    AddTranslation "ru", "core.letter.fallback.date", "Дата: "
    AddTranslation "en", "core.letter.fallback.date", "Date: "
    AddTranslation "ru", "core.letter.attachment_prefix", "Приложение: "
    AddTranslation "en", "core.letter.attachment_prefix", "Attachment: "
    AddTranslation "ru", "core.letter.text.confirmed", "направляем подтвержденные бухгалтерские документы в ваш адрес"
    AddTranslation "en", "core.letter.text.confirmed", "forwarding confirmed accounting documents to your address"
    AddTranslation "ru", "core.letter.text.own_confirmation", "направляем следующие документы в ваш адрес для подтверждения"
    AddTranslation "en", "core.letter.text.own_confirmation", "forwarding the following documents to your address for confirmation"
    AddTranslation "ru", "ribbon.msg.template_folder_saved", "Папка шаблонов сохранена:"
    AddTranslation "en", "ribbon.msg.template_folder_saved", "Template folder saved:"
    AddTranslation "ru", "ribbon.msg.output_folder_saved", "Папка для писем сохранена:"
    AddTranslation "en", "ribbon.msg.output_folder_saved", "Output folder saved:"
    AddTranslation "ru", "ribbon.msg.folder_select_error", "Ошибка выбора папки: "
    AddTranslation "en", "ribbon.msg.folder_select_error", "Folder selection error: "
    AddTranslation "ru", "ribbon.dialog.template_folder", "Выберите папку с шаблонами"
    AddTranslation "en", "ribbon.dialog.template_folder", "Select templates folder"
    AddTranslation "ru", "ribbon.dialog.output_folder", "Выберите папку для сформированных писем"
    AddTranslation "en", "ribbon.dialog.output_folder", "Select output folder"
    AddTranslation "ru", "ribbon.about.title", "О программе"
    AddTranslation "en", "ribbon.about.title", "About"
    AddTranslation "ru", "ribbon.about.name", "CreateLetter"
    AddTranslation "en", "ribbon.about.name", "CreateLetter"
    AddTranslation "ru", "ribbon.about.templates_folder", "Папка шаблонов: "
    AddTranslation "en", "ribbon.about.templates_folder", "Templates folder: "
    AddTranslation "ru", "ribbon.about.output_folder", "Папка выгрузки писем: "
    AddTranslation "en", "ribbon.about.output_folder", "Output folder: "
    AddTranslation "ru", "ribbon.about.open_form_hint", "Используйте ленту Excel для открытия формы и настройки папок."
    AddTranslation "en", "ribbon.about.open_form_hint", "Use the Excel ribbon to open the form and configure folders."
    AddTranslation "ru", "ribbon.msg.folder_unavailable", "Путь недоступен, использован путь книги:"
    AddTranslation "en", "ribbon.msg.folder_unavailable", "Configured path is unavailable, using workbook path:"
    AddTranslation "ru", "core.date.unknown_month", "неизвестный_месяц"
    AddTranslation "en", "core.date.unknown_month", "unknown_month"
    AddTranslation "ru", "core.form.open_creator_error", "Не удалось открыть форму создания письма: "
    AddTranslation "en", "core.form.open_creator_error", "Failed to open letter creation form: "
    AddTranslation "ru", "backup.msg.created_success", "Резервная копия создана успешно!"
    AddTranslation "en", "backup.msg.created_success", "Backup created successfully!"
    AddTranslation "ru", "backup.msg.create_error", "Ошибка при создании резервной копии: "
    AddTranslation "en", "backup.msg.create_error", "Error creating backup: "
    AddTranslation "ru", "backup.msg.folder_not_found", "Папка резервных копий не найдена."
    AddTranslation "en", "backup.msg.folder_not_found", "Backup folder not found."
    AddTranslation "ru", "backup.msg.list_title", "СПИСОК РЕЗЕРВНЫХ КОПИЙ:"
    AddTranslation "en", "backup.msg.list_title", "BACKUP LIST:"
    AddTranslation "ru", "backup.msg.none_found", "Резервные копии не найдены."
    AddTranslation "en", "backup.msg.none_found", "No backups found."
    AddTranslation "ru", "backup.msg.found_count", "Найдено резервных копий: "
    AddTranslation "en", "backup.msg.found_count", "Backups found: "
    AddTranslation "ru", "backup.msg.date_label", "  Дата: "
    AddTranslation "en", "backup.msg.date_label", "  Date: "
    AddTranslation "ru", "backup.msg.size_label", "  Размер: "
    AddTranslation "en", "backup.msg.size_label", "  Size: "
    AddTranslation "ru", "backup.title.information", "Информация о резервных копиях"
    AddTranslation "en", "backup.title.information", "Backup Information"
    AddTranslation "ru", "backup.title.restore", "Восстановление из резервной копии"
    AddTranslation "en", "backup.title.restore", "Restore from Backup"
    AddTranslation "ru", "backup.msg.restore_instructions", "Чтобы восстановить файл из резервной копии:"
    AddTranslation "en", "backup.msg.restore_instructions", "To restore from a backup:"
    AddTranslation "ru", "backup.msg.restore_step_1", "1. Закройте текущую книгу"
    AddTranslation "en", "backup.msg.restore_step_1", "1. Close the current workbook"
    AddTranslation "ru", "backup.msg.restore_step_2", "2. Откройте папку: "
    AddTranslation "en", "backup.msg.restore_step_2", "2. Open the folder: "
    AddTranslation "ru", "backup.msg.restore_step_3", "3. Выберите нужную резервную копию"
    AddTranslation "en", "backup.msg.restore_step_3", "3. Select the desired backup copy"
    AddTranslation "ru", "backup.msg.restore_step_4", "4. Переименуйте и откройте восстановленную книгу"
    AddTranslation "en", "backup.msg.restore_step_4", "4. Rename and open the restored workbook"
    AddTranslation "ru", "backup.msg.settings_saved", "Настройки резервного копирования сохранены:"
    AddTranslation "en", "backup.msg.settings_saved", "Backup settings saved:"
    AddTranslation "ru", "backup.msg.automatic_backup", "Автоматическое резервное копирование: "
    AddTranslation "en", "backup.msg.automatic_backup", "Automatic backup: "
    AddTranslation "ru", "backup.msg.keep_copies", "Хранить копии: "
    AddTranslation "en", "backup.msg.keep_copies", "Keep copies for: "
    AddTranslation "ru", "backup.label.enabled", "Включено"
    AddTranslation "en", "backup.label.enabled", "Enabled"
    AddTranslation "ru", "backup.label.disabled", "Отключено"
    AddTranslation "en", "backup.label.disabled", "Disabled"
    AddTranslation "ru", "backup.msg.settings_title", "НАСТРОЙКИ РЕЗЕРВНОГО КОПИРОВАНИЯ:"
    AddTranslation "en", "backup.msg.settings_title", "BACKUP SETTINGS:"
    AddTranslation "ru", "backup.msg.retention_period", "Период хранения: "
    AddTranslation "en", "backup.msg.retention_period", "Retention period: "
    AddTranslation "ru", "backup.msg.last_backup", "Последняя резервная копия: "
    AddTranslation "en", "backup.msg.last_backup", "Last backup: "
    AddTranslation "ru", "backup.msg.backup_folder", "Папка резервных копий: "
    AddTranslation "en", "backup.msg.backup_folder", "Backup folder: "
    AddTranslation "ru", "backup.label.never", "Никогда"
    AddTranslation "en", "backup.label.never", "Never"
    AddTranslation "ru", "snapshot.title.vba_snapshot", "Снимок VBA"
    AddTranslation "en", "snapshot.title.vba_snapshot", "VBA Snapshot"
    AddTranslation "ru", "snapshot.msg.vba_created_success", "Снимок VBA создан успешно!"
    AddTranslation "en", "snapshot.msg.vba_created_success", "VBA snapshot created successfully!"
    AddTranslation "ru", "snapshot.msg.folder_label", "Папка: "
    AddTranslation "en", "snapshot.msg.folder_label", "Folder: "
    AddTranslation "ru", "snapshot.msg.exported_components", "Экспортировано компонентов: "
    AddTranslation "en", "snapshot.msg.exported_components", "Exported components: "
    AddTranslation "ru", "snapshot.title.confirm_restore", "Подтвердите восстановление"
    AddTranslation "en", "snapshot.title.confirm_restore", "Confirm Restore"
    AddTranslation "ru", "snapshot.msg.restore_warning", "Внимание!" & vbCrLf & "Эта операция удалит текущие VBA-модули и восстановит их из выбранного снимка." & vbCrLf & "Продолжить?"
    AddTranslation "en", "snapshot.msg.restore_warning", "Warning!" & vbCrLf & "This operation removes the current VBA modules and restores them from the selected snapshot." & vbCrLf & "Do you want to continue?"
    AddTranslation "ru", "snapshot.msg.restore_complete", "Снимок успешно восстановлен!"
    AddTranslation "en", "snapshot.msg.restore_complete", "Snapshot restored successfully!"
    AddTranslation "ru", "snapshot.title.restore_complete", "Восстановление завершено"
    AddTranslation "en", "snapshot.title.restore_complete", "Restore Complete"
    AddTranslation "ru", "snapshot.msg.folder_not_found", "Папка снимков не найдена: "
    AddTranslation "en", "snapshot.msg.folder_not_found", "Snapshots folder not found: "
    AddTranslation "ru", "snapshot.prompt.select_folder", "Введите имя папки снимка:" & vbCrLf & "Доступный корень: "
    AddTranslation "en", "snapshot.prompt.select_folder", "Enter the snapshot folder name:" & vbCrLf & "Available root: "
    AddTranslation "ru", "snapshot.title.select_folder", "Выбор снимка"
    AddTranslation "en", "snapshot.title.select_folder", "Select Snapshot"
    AddTranslation "ru", "snapshot.prompt.default_folder", "Snapshot_"
    AddTranslation "en", "snapshot.prompt.default_folder", "Snapshot_"
    AddTranslation "ru", "snapshot.msg.version_tags_inserted", "Теги версии вставлены во все стандартные модули."
    AddTranslation "en", "snapshot.msg.version_tags_inserted", "Version tags inserted into all standard modules."
    AddTranslation "ru", "snapshot.title.workbook_snapshot", "Снимок книги"
    AddTranslation "en", "snapshot.title.workbook_snapshot", "Workbook Snapshot"
    AddTranslation "ru", "snapshot.prompt.workbook_label", "Введите краткую метку снимка:"
    AddTranslation "en", "snapshot.prompt.workbook_label", "Enter a short snapshot label:"
    AddTranslation "ru", "snapshot.prompt.workbook_default", "manual_snapshot"
    AddTranslation "en", "snapshot.prompt.workbook_default", "manual_snapshot"
    AddTranslation "ru", "snapshot.msg.workbook_created", "Снимок книги создан!"
    AddTranslation "en", "snapshot.msg.workbook_created", "Workbook snapshot created!"
    AddTranslation "ru", "snapshot.msg.file_label", "Файл: "
    AddTranslation "en", "snapshot.msg.file_label", "File: "
    AddTranslation "ru", "bootstrap.msg.structure_created", "Структура листов создана успешно!"
    AddTranslation "en", "bootstrap.msg.structure_created", "Sheets structure created successfully!"
    AddTranslation "ru", "bootstrap.msg.structure_create_error", "Ошибка создания структуры листов: "
    AddTranslation "en", "bootstrap.msg.structure_create_error", "Error creating sheets structure: "
    AddTranslation "ru", "bootstrap.msg.reset_confirm", "Вы уверены, что хотите сбросить все данные?"
    AddTranslation "en", "bootstrap.msg.reset_confirm", "Are you sure you want to reset all data?"
    AddTranslation "ru", "audit.msg.log_opened", "Журнал аудита открыт. При необходимости снова скройте лист после просмотра."
    AddTranslation "en", "audit.msg.log_opened", "Audit log opened. Hide the sheet after review if needed."
    AddTranslation "ru", "audit.msg.report_error", "Ошибка при формировании отчета аудита: "
    AddTranslation "en", "audit.msg.report_error", "Error generating audit report: "
    AddTranslation "ru", "audit.title.system_statistics", "Статистика системы"
    AddTranslation "en", "audit.title.system_statistics", "System Statistics"
    AddTranslation "ru", "audit.msg.system_statistics_header", "СТАТИСТИКА ИСПОЛЬЗОВАНИЯ СИСТЕМЫ"
    AddTranslation "en", "audit.msg.system_statistics_header", "SYSTEM USAGE STATISTICS"
    AddTranslation "ru", "audit.msg.total_sessions", "Всего сессий: "
    AddTranslation "en", "audit.msg.total_sessions", "Total sessions: "
    AddTranslation "ru", "audit.msg.letters_created", "Создано писем: "
    AddTranslation "en", "audit.msg.letters_created", "Letters created: "
    AddTranslation "ru", "audit.msg.unique_users", "Уникальных пользователей: "
    AddTranslation "en", "audit.msg.unique_users", "Unique users: "
    AddTranslation "ru", "audit.msg.top_users", "ТОП ПОЛЬЗОВАТЕЛЕЙ:"
    AddTranslation "en", "audit.msg.top_users", "TOP USERS:"
    AddTranslation "ru", "audit.title.admin_panel", "Панель администратора"
    AddTranslation "en", "audit.title.admin_panel", "Administrator Panel"
    AddTranslation "ru", "audit.msg.admin_prompt", "Выберите действие:"
    AddTranslation "en", "audit.msg.admin_prompt", "Select action:"
    AddTranslation "ru", "audit.msg.admin_option_1", "1 - Показать журнал аудита"
    AddTranslation "en", "audit.msg.admin_option_1", "1 - Show audit log"
    AddTranslation "ru", "audit.msg.admin_option_2", "2 - Статистика использования"
    AddTranslation "en", "audit.msg.admin_option_2", "2 - Usage statistics"
    AddTranslation "ru", "audit.msg.admin_option_3", "3 - Отчет за 30 дней"
    AddTranslation "en", "audit.msg.admin_option_3", "3 - 30-day report"
    AddTranslation "ru", "audit.msg.admin_option_4", "4 - Отчет за 7 дней"
    AddTranslation "en", "audit.msg.admin_option_4", "4 - 7-day report"
    AddTranslation "ru", "audit.msg.cancelled", "Отменено"
    AddTranslation "en", "audit.msg.cancelled", "Cancelled"
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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub
    
    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Dim entryKey As String
        Dim entryRu As String
        Dim entryEn As String
        
        entryKey = NormalizeLocalizationKey(CStr(ws.Cells(rowIndex, 1).value))
        If Len(entryKey) = 0 Then GoTo NextRow
        
        entryRu = CStr(ws.Cells(rowIndex, 2).value)
        entryEn = CStr(ws.Cells(rowIndex, 3).value)
        
        If Len(entryRu) > 0 Then AddTranslation "ru", entryKey, entryRu
        If Len(entryEn) > 0 Then AddTranslation "en", entryKey, entryEn
        
NextRow:
    Next rowIndex
    
    Exit Sub
    
LoadError:
    Debug.Print "Localization load skipped: " & Err.description
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


