Attribute VB_Name = "ModuleWordInterop"
' ======================================================================
' Module: ModuleWordInterop
' Author: CreateLetter contributors
' Purpose: Explicit Word session lifecycle and document generation helpers
' Version: 1.1.2 - 29.03.2026
' ======================================================================

Option Explicit

Private g_WordApp As Object
Private g_WordAppOwned As Boolean

Public Function TryAttachRunningWordApplication(ByRef wordApp As Object) As Boolean
    On Error GoTo LookupFailed

    Set wordApp = GetObject(, "Word.Application")
    TryAttachRunningWordApplication = Not wordApp Is Nothing
    Exit Function

LookupFailed:
    Set wordApp = Nothing
    TryAttachRunningWordApplication = False
End Function

Public Function CreateOwnedWordApplication() As Object
    On Error GoTo CreateFailed

    Set CreateOwnedWordApplication = CreateObject("Word.Application")
    g_WordAppOwned = True
    Exit Function

CreateFailed:
    Set CreateOwnedWordApplication = Nothing
    g_WordAppOwned = False
    Err.Raise Err.Number, "CreateOwnedWordApplication", Err.description
End Function

Public Function AcquireWordApplication() As Object
    On Error GoTo AcquireFailed

    If IsWordApplicationAlive(g_WordApp) Then
        Set AcquireWordApplication = g_WordApp
        Exit Function
    End If

    Set g_WordApp = Nothing
    g_WordAppOwned = False

    If TryAttachRunningWordApplication(g_WordApp) Then
        Set AcquireWordApplication = g_WordApp
        Exit Function
    End If

    Set g_WordApp = CreateOwnedWordApplication()
    Set AcquireWordApplication = g_WordApp
    Exit Function

AcquireFailed:
    Set g_WordApp = Nothing
    g_WordAppOwned = False
    Err.Raise Err.Number, "AcquireWordApplication", Err.description
End Function

Public Sub ReleaseWordApplication(Optional closeDocuments As Boolean = False)
    On Error GoTo ReleaseFailed

    If g_WordApp Is Nothing Then Exit Sub

    If g_WordAppOwned Then
        If closeDocuments Then
            g_WordApp.Quit False
        Else
            g_WordApp.Visible = True
        End If
    End If

    Set g_WordApp = Nothing
    g_WordAppOwned = False
    Exit Sub

ReleaseFailed:
    Set g_WordApp = Nothing
    g_WordAppOwned = False
End Sub

Public Sub ResetStaleWordApplication()
    Set g_WordApp = Nothing
    g_WordAppOwned = False
End Sub

Public Function GetWordApplicationState() As String
    If g_WordApp Is Nothing Then
        GetWordApplicationState = "empty"
    ElseIf IsWordApplicationAlive(g_WordApp) Then
        If g_WordAppOwned Then
            GetWordApplicationState = "owned"
        Else
            GetWordApplicationState = "reused"
        End If
    Else
        GetWordApplicationState = "stale"
    End If
End Function

Public Function WarmUpWordApplication() As Boolean
    On Error GoTo WarmUpFailed

    Dim wordApp As Object
    Set wordApp = AcquireWordApplication()
    wordApp.Visible = True

    WarmUpWordApplication = Not wordApp Is Nothing
    Exit Function

WarmUpFailed:
    WarmUpWordApplication = False
End Function

Public Sub WordInteropCreateLetterDocument(Addressee As String, addressArray As Variant, letterNumber As String, letterDateRaw As String, Executor As String, documentType As String, useAlternateTemplate As Boolean, documentsList As Collection)
    Dim wordApp As Object
    Dim wordDoc As Object

    On Error GoTo ErrorHandler

    Set wordApp = AcquireWordApplication()
    wordApp.Visible = True

    Dim templatePath As String
    templatePath = GetLetterTemplatePathInternal(useAlternateTemplate)

    If dir$(templatePath) <> "" Then
        Set wordDoc = wordApp.documents.Open(templatePath)
        If Not wordDoc Is Nothing Then
            WordInteropFillWordTemplateData wordDoc, Addressee, addressArray, letterNumber, letterDateRaw, Executor, documentType, documentsList
            GoTo SaveDocument
        End If
    End If

    Set wordDoc = wordApp.documents.Add
    WordInteropCreateLetterDocumentFromScratch wordDoc, Addressee, addressArray, letterNumber, letterDateRaw, Executor, documentType, documentsList

SaveDocument:
    Dim fileName As String
    fileName = GenerateFileNameWithExecutor(IIf(Len(Trim$(Addressee)) = 0, t("core.letter.default_file_name", "Письмо"), Addressee), letterNumber, Executor)

    wordDoc.SaveAs fileName
    Debug.Print "File saved: " & fileName

    Dim saveWorkbookError As String
    If TrySaveCurrentWorkbook(saveWorkbookError) Then
        Debug.Print "Excel workbook saved"
    Else
        Debug.Print "Excel workbook save failed: " & saveWorkbookError
        MsgBox t("core.letter.warning.workbook_not_saved", "Файл письма создан, но книгу не удалось сохранить: ") & saveWorkbookError, vbExclamation
    End If

    wordApp.Visible = True
    wordDoc.Activate
    ReleaseWordApplication False

    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox t("core.letter.error.create_document", "Ошибка при создании письма: ") & Err.description, vbCritical
    If Not wordDoc Is Nothing Then
        wordDoc.Close False
    End If
    If Not IsWordApplicationAlive(wordApp) Then
        ResetStaleWordApplication
    Else
        ReleaseWordApplication False
    End If
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Public Sub WordInteropFillWordTemplateData(wordDoc As Object, addresseeText As String, addressArray As Variant, numberText As String, rawDateText As String, executorText As String, documentType As String, documentsList As Collection)
    On Error GoTo TemplateError

    Dim addressText As String
    Dim dateText As String
    Dim phoneText As String
    Dim letterText As String

    addressText = FormatRecipientAddress(addressArray)
    dateText = FormatLetterDate(rawDateText)
    phoneText = GetExecutorPhone(executorText)
    letterText = GetDocumentTypeText(documentType)

    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderRecipientName, LegacyTemplatePlaceholderRecipientName, addresseeText
    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderRecipientAddress, LegacyTemplatePlaceholderRecipientAddress, addressText
    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderOutgoingNumber, LegacyTemplatePlaceholderOutgoingNumber, numberText
    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderOutgoingDate, LegacyTemplatePlaceholderOutgoingDate, dateText
    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderExecutorName, LegacyTemplatePlaceholderExecutorName, executorText
    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderExecutorPhone, LegacyTemplatePlaceholderExecutorPhone, phoneText
    ReplaceTemplatePlaceholderVariants wordDoc, TemplatePlaceholderLetterText, LegacyTemplatePlaceholderLetterText, letterText

    WordInteropReplaceAttachmentsInTemplateWithFontAndSum wordDoc, documentsList, 10
    Exit Sub

TemplateError:
    MsgBox t("core.letter.error.template_fill", "Ошибка заполнения шаблона: ") & Err.description, vbCritical
End Sub

Public Sub WordInteropCreateLetterDocumentFromScratch(wordDoc As Object, addresseeText As String, addressArray As Variant, numberText As String, rawDateText As String, executorText As String, documentType As String, documentsList As Collection)
    On Error GoTo ScratchError

    Dim content As String
    Dim addressText As String
    Dim letterText As String
    Dim dateText As String

    addressText = FormatRecipientAddress(addressArray)
    letterText = GetDocumentTypeText(documentType)
    dateText = FormatLetterDate(rawDateText)

    content = t("core.letter.fallback.to_commander", "Командиру войсковой части ") & addresseeText & vbCrLf & vbCrLf
    content = content & addressText & vbCrLf & vbCrLf & vbCrLf
    content = content & letterText & vbCrLf & vbCrLf
    content = content & t("core.letter.fallback.executor", "Исполнитель: ") & executorText & vbCrLf
    content = content & t("core.letter.fallback.phone", "Телефон: ") & GetExecutorPhone(executorText) & vbCrLf
    content = content & t("core.letter.fallback.ref_no", "Исх. №: ") & numberText & vbCrLf
    content = content & t("core.letter.fallback.date", "Дата: ") & dateText & vbCrLf & vbCrLf

    wordDoc.content.Text = content
    WordInteropAppendAttachmentsToDocumentWithFontAndSum wordDoc, documentsList, 10
    Exit Sub

ScratchError:
    MsgBox t("core.letter.error.create_fallback", "Ошибка при создании текста письма: ") & Err.description, vbCritical
End Sub

Public Sub WordInteropReplaceAttachmentsInTemplateWithFontAndSum(wordDoc As Object, documentsList As Collection, fontSize As Integer)
    On Error GoTo ReplaceError

    Dim rng As Object
    Set rng = wordDoc.content

    With rng.Find
        .ClearFormatting
        .Forward = True
        .Wrap = 1
        .Text = ResolveAttachmentsPlaceholderText(wordDoc)

        If Len(.Text) > 0 And .Execute Then
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

            WordInteropFormatAttachmentsInWord attachmentRange, fontSize
        End If
    End With

    Exit Sub

ReplaceError:
    Err.Raise Err.Number, "WordInteropReplaceAttachmentsInTemplateWithFontAndSum", Err.description
End Sub

Public Sub WordInteropAppendAttachmentsToDocumentWithFontAndSum(wordDoc As Object, documentsList As Collection, fontSize As Integer)
    On Error GoTo AppendError

    Dim rng As Object
    Set rng = wordDoc.content
    rng.Collapse 0

    rng.InsertAfter t("core.letter.attachment_prefix", "Приложение: ")

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

    WordInteropFormatAttachmentsInWord attachmentRange, fontSize
    rng.InsertAfter vbCrLf & vbCrLf
    Exit Sub

AppendError:
    Err.Raise Err.Number, "WordInteropAppendAttachmentsToDocumentWithFontAndSum", Err.description
End Sub

Public Sub WordInteropSafeReplaceInWord(wordDoc As Object, findText As String, replaceText As String)
    On Error GoTo ReplaceError

    If Len(replaceText) > 180 Then
        Dim fragments As Collection
        Set fragments = SplitStringToFragments(replaceText, 180)
        WordInteropSafeReplaceInWordWithFragments wordDoc, findText, fragments
    Else
        With wordDoc.content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .Text = findText
            .Replacement.Text = replaceText
            .Execute Replace:=2
        End With

        If findText = TemplatePlaceholderExecutorPhone Then
            Dim rng As Object
            Set rng = wordDoc.content
            With rng.Find
                .Text = replaceText
                If .Execute Then
                    ApplyWordRangeFontFormatting rng, "Times New Roman", 12
                    rng.Font.Color = RGB(0, 0, 0)
                End If
            End With
        End If
    End If

    Exit Sub

ReplaceError:
    If Not TryFallbackReplaceWordContent(wordDoc, findText, replaceText) Then
        Err.Raise Err.Number, "WordInteropSafeReplaceInWord", Err.description
    End If
End Sub

Public Sub WordInteropSafeReplaceInWordWithFragments(wordDoc As Object, findText As String, fragments As Collection)
    On Error GoTo ReplaceError

    Dim rng As Object
    Set rng = wordDoc.content

    With rng.Find
        .ClearFormatting
        .Forward = True
        .Wrap = 1
        .Text = findText

        If .Execute Then
            rng.Delete

            Dim i As Long
            Dim fullText As String
            For i = 1 To fragments.count
                If i > 1 Then fullText = fullText & " "
                fullText = fullText & CStr(fragments(i))
            Next i

            rng.InsertAfter fullText

            Dim insertedRange As Object
            Set insertedRange = wordDoc.Range(rng.Start, rng.Start + Len(fullText))
            ApplyWordRangeFontFormatting insertedRange, "Times New Roman", 12
        End If
    End With

    Exit Sub

ReplaceError:
    Debug.Print "WordInteropSafeReplaceInWordWithFragments error: " & Err.description
End Sub

Public Sub WordInteropFormatAttachmentsInWord(rng As Object, Optional fontSize As Integer = 10)
    On Error GoTo FormatError

    ApplyWordRangeFontFormatting rng, rng.Font.Name, fontSize
    Exit Sub

FormatError:
    Err.Raise Err.Number, "WordInteropFormatAttachmentsInWord", Err.description
End Sub

Private Function IsWordApplicationAlive(wordApp As Object) As Boolean
    On Error GoTo NotAlive

    If wordApp Is Nothing Then Exit Function
    Dim visibleState As Boolean
    visibleState = CBool(wordApp.Visible)
    IsWordApplicationAlive = True
    Exit Function

NotAlive:
    IsWordApplicationAlive = False
End Function

Private Function TrySaveCurrentWorkbook(ByRef errorMessage As String) As Boolean
    On Error GoTo SaveFailed

    ThisWorkbook.Save
    errorMessage = ""
    TrySaveCurrentWorkbook = True
    Exit Function

SaveFailed:
    errorMessage = Err.description
    TrySaveCurrentWorkbook = False
End Function

Private Sub ApplyWordRangeFontFormatting(targetRange As Object, fontName As String, fontSize As Integer)
    With targetRange
        .Font.Name = fontName
        .Font.Size = fontSize
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.LineSpacing = fontSize + 2
    End With
End Sub

Private Function TryFallbackReplaceWordContent(wordDoc As Object, findText As String, replaceText As String) As Boolean
    On Error GoTo FallbackFailed

    wordDoc.content.Text = Replace(wordDoc.content.Text, findText, Left$(replaceText, 200))
    TryFallbackReplaceWordContent = True
    Exit Function

FallbackFailed:
    TryFallbackReplaceWordContent = False
End Function

Private Function SplitStringToFragments(inputString As String, maxLength As Integer) As Collection
    Set SplitStringToFragments = New Collection

    If Len(inputString) = 0 Then Exit Function

    Dim currentPos As Long
    currentPos = 1

    While currentPos <= Len(inputString)
        If currentPos + maxLength - 1 > Len(inputString) Then
            SplitStringToFragments.Add Mid$(inputString, currentPos)
            Exit Function
        End If

        Dim fragment As String
        fragment = Mid$(inputString, currentPos, maxLength)

        Dim breakPos As Long
        breakPos = FindBestBreakPosition(fragment)

        If breakPos > 0 And breakPos < maxLength Then
            fragment = Mid$(inputString, currentPos, breakPos)
            currentPos = currentPos + breakPos + 1

            While currentPos <= Len(inputString) And Mid$(inputString, currentPos, 1) = " "
                currentPos = currentPos + 1
            Wend
        Else
            currentPos = currentPos + maxLength
        End If

        SplitStringToFragments.Add Trim$(fragment)
    Wend
End Function

Private Function FindBestBreakPosition(textFragment As String) As Long
    Dim i As Long
    Dim testPos As Long

    For i = Len(textFragment) To Len(textFragment) \ 2 Step -1
        Dim currentChar As String
        currentChar = Mid$(textFragment, i, 1)

        If currentChar = "." And i < Len(textFragment) Then
            If Mid$(textFragment, i + 1, 1) = " " Then
                FindBestBreakPosition = i
                Exit Function
            End If
        End If

        If currentChar = "," And i < Len(textFragment) Then
            If Mid$(textFragment, i + 1, 1) = " " Then
                testPos = i
            End If
        End If

        If currentChar = ":" And i < Len(textFragment) Then
            If Mid$(textFragment, i + 1, 1) = " " Then
                If testPos = 0 Then testPos = i
            End If
        End If

        If currentChar = " " And testPos = 0 Then
            testPos = i - 1
        End If
    Next i

    FindBestBreakPosition = testPos
End Function

Private Function GetLetterTemplatePathInternal(useAlternateTemplate As Boolean) As String
    Dim templateFolder As String
    templateFolder = GetConfiguredTemplateFolderPath()

    If useAlternateTemplate Then
        GetLetterTemplatePathInternal = templateFolder & "\" & LetterTemplateFileNameFOU
    Else
        GetLetterTemplatePathInternal = templateFolder & "\" & LetterTemplateFileNameRegular
    End If
End Function

Private Sub ReplaceTemplatePlaceholderVariants(wordDoc As Object, primaryPlaceholder As String, legacyPlaceholder As String, replaceText As String)
    WordInteropSafeReplaceInWord wordDoc, primaryPlaceholder, replaceText

    If Len(Trim$(legacyPlaceholder)) > 0 Then
        WordInteropSafeReplaceInWord wordDoc, legacyPlaceholder, replaceText
    End If
End Sub

Private Function ResolveAttachmentsPlaceholderText(wordDoc As Object) As String
    On Error GoTo LookupFailed

    Dim documentText As String
    documentText = CStr(wordDoc.content.Text)

    If InStr(1, documentText, TemplatePlaceholderAttachmentsList, vbTextCompare) > 0 Then
        ResolveAttachmentsPlaceholderText = TemplatePlaceholderAttachmentsList
        Exit Function
    End If

    If InStr(1, documentText, LegacyTemplatePlaceholderAttachmentsList, vbTextCompare) > 0 Then
        ResolveAttachmentsPlaceholderText = LegacyTemplatePlaceholderAttachmentsList
        Exit Function
    End If

LookupFailed:
    ResolveAttachmentsPlaceholderText = TemplatePlaceholderAttachmentsList
End Function