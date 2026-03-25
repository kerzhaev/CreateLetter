Attribute VB_Name = "ModuleLocalization"
' ======================================================================
' Module: ModuleLocalization
' Purpose: Provide a safe localization foundation before UI text migration
' Version: 1.0.0
' Notes:
' - This module does not change workbook schema.
' - Workbook sheet-based localization is optional and loaded only if present.
' - Built-in Russian values keep current user-facing language stable.
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
    ' Built-in defaults keep the current runtime language stable until
    ' workbook-backed localization data is introduced in a later stage.
    AddTranslation "ru", "app.language.name", "Русский"
    AddTranslation "en", "app.language.name", "English"
    AddTranslation "ru", "common.ok", "ОК"
    AddTranslation "ru", "common.cancel", "Отмена"
    AddTranslation "ru", "common.yes", "Да"
    AddTranslation "ru", "common.no", "Нет"
    AddTranslation "ru", "dialog.cancel_letter_creation", "Отменить создание письма?"
    AddTranslation "ru", "dialog.discard_unsaved_documents", "Несохраненные документы будут потеряны. Закрыть?"
    AddTranslation "ru", "form.letter_creator.title", "Формирование писем"
    AddTranslation "ru", "form.letter_history.title", "История отправленных писем"
    AddTranslation "ru", "form.letter_creator.tip.document_sum", "Сумма документа в рублях (необязательно). Например: 125000"
    AddTranslation "ru", "form.letter_creator.progress.page", "Шаг"
    AddTranslation "ru", "form.letter_creator.attachments_count", "Выбрано документов:"
    AddTranslation "ru", "form.letter_creator.caption.edit_address", "Изменить адрес"
    AddTranslation "ru", "form.letter_creator.tip.edit_address", "Редактировать выбранный адрес"
    AddTranslation "ru", "form.letter_creator.caption.delete_address", "Удалить адрес"
    AddTranslation "ru", "form.letter_creator.tip.delete_address", "Удалить выбранный адрес"
    AddTranslation "ru", "form.letter_creator.tip.phone", "Телефон адресата (формат: 8-xxx-xxx-xx-xx)"
    AddTranslation "ru", "form.letter_creator.caption.letter_history", "История писем"
    AddTranslation "ru", "form.letter_creator.tip.letter_history", "Открыть форму истории отправленных писем"
    AddTranslation "ru", "form.letter_creator.tip.selected_attachments", "Для просмотра полного названия наведите на элемент"
    AddTranslation "ru", "form.letter_creator.tip.address_search", "Введите часть наименования для поиска адресата"
    AddTranslation "ru", "form.letter_creator.tip.letter_number", "Введите номер после 7/ (например: 125 > получится 7/125)"
    AddTranslation "ru", "form.letter_creator.tip.letter_date", "Формат: дд.мм.гггг"
    AddTranslation "ru", "form.letter_creator.caption.next", "Далее >"
    AddTranslation "ru", "form.letter_creator.caption.create_letter", "СОЗДАТЬ ПИСЬМО"
    AddTranslation "ru", "status.ready", "Готово"
    AddTranslation "ru", "status.searching", "Поиск..."
    AddTranslation "ru", "error.generic", "Произошла ошибка."
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
