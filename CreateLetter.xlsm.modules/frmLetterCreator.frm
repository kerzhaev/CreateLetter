VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLetterCreator 
   Caption         =   "Letter Builder v1.6.18"
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16320
   OleObjectBlob   =   "frmLetterCreator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLetterCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

















' ======================================================================

' Form    : frmLetterCreator v1.6.18 - Thin-shell MultiPage wizard with workbook-backed localization, grouped address search, and explicit startup focus

' Version : 1.6.18 - 17.04.2026

' Author  : CreateLetter contributors

' Purpose : UI orchestration for letter creation, address entry, attachments, summary flow, and schema-safe bindings

' ======================================================================



Option Explicit



'------------------------------------------------------------

'  GLOBAL VARIABLES

'------------------------------------------------------------

Private Const TOTAL_PAGES As Integer = 4

Private Const ADDRESS_GROUP_LABEL_NAME As String = "lblAddressGroup"

Private Const ADDRESS_GROUP_TEXTBOX_NAME As String = "txtAddressGroup"

Public selectedAddressRow As Long

Private documentsList As Collection

Private contextMenuSelectedIndex As Integer

Private currentAddressSearchResults As Collection

Private isClosingForm As Boolean

Private skipNextAddressAutoUpdate As Boolean

Private pendingInitialSearchFocus As Boolean



'------------------------------------------------------------

'  FORM INITIALIZATION

'------------------------------------------------------------

Private Sub UserForm_Initialize()

    Set documentsList = New Collection

    Set currentAddressSearchResults = New Collection

    contextMenuSelectedIndex = -1

    isClosingForm = False

    skipNextAddressAutoUpdate = False

    pendingInitialSearchFocus = True

    

    ClearAddressCache

    EnsureAddressGroupControls

    ConfigureFormAppearance

    ApplyLocalizedStaticCaptions

    InitializeControlValues

    LoadExecutorsList

    

    ConfigureMultilineTextBoxes

    ConfigureDocumentSumField

    

    ClearDocumentFields

    InitializeProgressInfo

    

    SwitchToPage 0

End Sub

Private Sub UserForm_Activate()

    If pendingInitialSearchFocus Then

        ApplyInitialSearchFocus

    End If

End Sub

Public Sub ApplyInitialSearchFocus()

    pendingInitialSearchFocus = False

    On Error Resume Next

    If Not mpgWizard Is Nothing Then mpgWizard.value = 0

    Me.Repaint
    DoEvents
    SafeSetFocus "txtAddressSearch"

    On Error GoTo 0

End Sub



Private Sub ConfigureDocumentSumField()

    On Error Resume Next

    

    If Not txtDocumentSum Is Nothing Then

        With txtDocumentSum

            .Font.Name = "Segoe UI"

            .Font.Size = 10

            .ControlTipText = t("form.letter_creator.tip.document_sum", "Document amount in rubles (optional). Example: 125000")

            .value = ""

            .backColor = RGB(255, 255, 255)

        End With

        Debug.Print "Document sum field configured"

    End If

    

    On Error GoTo 0

End Sub



Private Sub InitializeProgressInfo()

    lblProgressInfo.Caption = BuildCreatorProgressCaption(1, TOTAL_PAGES)

    lblAttachmentsCount.Caption = BuildCreatorSelectedDocumentsCaption(0)

End Sub



'------------------------------------------------------------

'  CONTROL VALUES INITIALIZATION

'------------------------------------------------------------

Private Sub InitializeControlValues()

    On Error Resume Next

    

    Me.Controls("txtLetterDate").value = Format(Date, "dd.mm.yyyy")

    Me.Controls("txtLetterNumber").value = "7/"

    

    PopulateDocumentTypeOptions Me.Controls("cmbDocumentType")

    PopulateLetterTypeOptions Me.Controls("cmbLetterType")

    

    selectedAddressRow = 0

    On Error GoTo 0

End Sub



'------------------------------------------------------------

'  FIELD CHANGE EVENTS

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

        currentValue = Trim(txtDocumentSum.value)

        

        If Len(currentValue) > 0 And Not IsNumeric(currentValue) Then

            txtDocumentSum.backColor = RGB(255, 240, 240)

        Else

            txtDocumentSum.backColor = RGB(255, 255, 255)

        End If

    End If

    

    On Error GoTo 0

End Sub



'------------------------------------------------------------

'  FOCUS LOST EVENTS

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

'  LETTER HISTORY BUTTON

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

        Debug.Print "History form is already open - activated"

    Else

        Load frmLetterHistory

        frmLetterHistory.Show vbModeless

        Debug.Print "History form opened modelessly"

    End If

    

    Exit Sub

    

HistoryError:

    MsgBox t("form.letter_creator.msg.history_open_error", "Error opening the history form: ") & Err.description, vbCritical

End Sub



'------------------------------------------------------------

'  CLEAR SEARCH BUTTON

'------------------------------------------------------------

Private Sub btnClearSearch_Click()

    On Error Resume Next

    

    Me.Controls("txtAddressSearch").value = ""

    Me.Controls("lstAddresses").Clear

    Set currentAddressSearchResults = New Collection

    

    ClearAllAddressFields

    ResetAddressFormState

    

    Me.Controls("txtAddressSearch").SetFocus

    

    Debug.Print "Full search and address fields clear performed"

    

    On Error GoTo 0

End Sub



Private Sub ClearAllAddressFields()

    On Error Resume Next

    

    Dim addressFields As Variant

    Dim i As Long

    Dim ctrl As control

    

    addressFields = Array("txtAddressee", "txtStreet", "txtCity", "txtDistrict", "txtRegion", "txtPostalCode", "txtAddresseePhone", ADDRESS_GROUP_TEXTBOX_NAME)

    

    For i = LBound(addressFields) To UBound(addressFields)

        Set ctrl = ResolveNamedControl(CStr(addressFields(i)))

        If Not ctrl Is Nothing Then

            ctrl.value = ""

            ctrl.backColor = RGB(255, 255, 255)

            Debug.Print "Cleared field: " & addressFields(i)

        End If

    Next i

    

    selectedAddressRow = 0

    Debug.Print "All address fields cleared"

    

    On Error GoTo 0

End Sub



Private Sub CheckRequiredFields()

    On Error Resume Next

    

    If Len(Trim(Me.Controls("txtCity").value)) = 0 Then

        Me.Controls("txtCity").backColor = RGB(255, 240, 240)

    Else

        Me.Controls("txtCity").backColor = RGB(240, 255, 240)

    End If

    

    If Len(Trim(Me.Controls("txtRegion").value)) = 0 Then

        Me.Controls("txtRegion").backColor = RGB(255, 240, 240)

    Else

        Me.Controls("txtRegion").backColor = RGB(240, 255, 240)

    End If

    

    If Len(Trim(Me.Controls("txtPostalCode").value)) = 0 Then

        Me.Controls("txtPostalCode").backColor = RGB(255, 240, 240)

    Else

        Me.Controls("txtPostalCode").backColor = RGB(240, 255, 240)

    End If

    

    If Me.Controls("cmbExecutor").listIndex < 0 Or Len(Trim(Me.Controls("cmbExecutor").value)) = 0 Then

        Me.Controls("cmbExecutor").backColor = RGB(255, 240, 240)

    Else

        Me.Controls("cmbExecutor").backColor = RGB(240, 255, 240)

    End If

    

    On Error GoTo 0

End Sub



'------------------------------------------------------------

'  APPEARANCE CONFIGURATION

'------------------------------------------------------------

Private Sub ConfigureFormAppearance()

    Me.Font.Name = "Segoe UI"

    Me.Font.Size = 10

    Me.Caption = t("form.letter_creator.title", "Letter Builder") & " v1.6.18"

    

    On Error Resume Next

    

    If Not lstSelectedAttachments Is Nothing Then

        lstSelectedAttachments.Font.Size = 9

        lstSelectedAttachments.ControlTipText = t("form.letter_creator.tip.selected_attachments", "Hover over the item to see the full name")

        lstSelectedAttachments.IntegralHeight = False

    End If

    

    If Not btnEditAddress Is Nothing Then

        btnEditAddress.Caption = t("form.letter_creator.caption.edit_address", "Edit address")

        btnEditAddress.ControlTipText = t("form.letter_creator.tip.edit_address", "Edit selected address")

        btnEditAddress.Enabled = False

    End If

    

    If Not btnDeleteAddress Is Nothing Then

        btnDeleteAddress.Caption = t("form.letter_creator.caption.delete_address", "Delete address")

        btnDeleteAddress.ControlTipText = t("form.letter_creator.tip.delete_address", "Delete selected address")

        btnDeleteAddress.Enabled = False

    End If

    

    If Not txtAddresseePhone Is Nothing Then

        txtAddresseePhone.ControlTipText = t("form.letter_creator.tip.phone", "Addressee phone (format: 8-xxx-xxx-xx-xx)")

        txtAddresseePhone.Enabled = True

        txtAddresseePhone.backColor = RGB(255, 255, 255)

    End If

    

    If Not btnLetterHistory Is Nothing Then

        With btnLetterHistory

            .Caption = t("form.letter_creator.caption.letter_history", "Letters History")

            .Font.Name = "Segoe UI"

            .Font.Size = 10

            .Font.Bold = True

            .backColor = RGB(156, 39, 176)

            .ForeColor = RGB(255, 255, 255)

            .ControlTipText = t("form.letter_creator.tip.letter_history", "Open sent letters history form")

        End With

    End If

    

    On Error GoTo 0

    

    txtAddressSearch.ControlTipText = t("form.letter_creator.tip.address_search", "Enter part of the name to search for the addressee")

    txtLetterNumber.ControlTipText = t("form.letter_creator.tip.letter_number", "Enter the number after 7/ (for example: 125 becomes 7/125)")

    txtLetterDate.ControlTipText = t("form.letter_creator.tip.letter_date", "Format: dd.mm.yyyy")

    SetResolvedControlTip ADDRESS_GROUP_TEXTBOX_NAME, GetAddressGroupTooltipText()

End Sub



Private Sub ApplyLocalizedStaticCaptions()

    On Error Resume Next



    SetLocalizedCaption "lblStep1", "form.letter_creator.label.stage", "Stage:"

    SetControlCaption "lblStep2", ""

    SetControlCaption "lblStep3", ""

    SetControlCaption "lblStep4", ""

    SetControlCaption "lblStep5", ""

    SetLocalizedCaption "lblCurrentAction", "form.letter_creator.label.current_action", "Current action"

    SetLocalizedCaption "Label1", "form.letter_creator.label.search_addressee", "Search existing addressee"

    SetLocalizedCaption "Label2", "form.letter_creator.label.city", "City"

    SetLocalizedCaption "Label3", "form.letter_creator.label.district", "District"

    SetLocalizedCaption "Label4", "form.letter_creator.label.region", "Region"

    SetLocalizedCaption "Label5", "form.letter_creator.label.postal_code", "Postal code"

    SetLocalizedCaption "Label6", "form.letter_creator.label.executor", "Executor"

    SetLocalizedCaption "Label7", "form.letter_creator.label.letter_date", "Letter date"

    SetLocalizedCaption "Label8", "form.letter_creator.label.letter_number", "Letter number"

    SetLocalizedCaption "Label9", "form.letter_creator.label.search_attachment", "Search attachment"

    SetLocalizedCaption "Label10", "form.letter_creator.label.selected_attachments", "Selected attachments"

    SetLocalizedCaption "Label11", "form.letter_creator.label.document_ownership", "Document type"

    SetLocalizedCaption "Label13", "form.letter_creator.label.date", "Date"

    SetLocalizedCaption "Label14", "form.letter_creator.label.copies", "Copies"

    SetLocalizedCaption "Label15", "form.letter_creator.label.sheets", "Sheets"

    SetLocalizedCaption "Label16", "form.letter_creator.label.found_addresses", "Found addresses"

    SetLocalizedCaption "Label17", "form.letter_creator.label.street_house", "Street, house"

    SetLocalizedCaption "Label18", "form.letter_creator.label.addressee", "Addressee"

    SetResolvedControlCaption ADDRESS_GROUP_LABEL_NAME, t("form.letter_creator.label.address_group", "Address group")

    SetLocalizedCaption "Label19", "form.letter_creator.label.available_attachments", "Available attachments"

    SetLocalizedCaption "Label20", "form.letter_creator.label.number", "Number"

    SetLocalizedCaption "Label21", "form.letter_creator.label.summary_addressee", "Addressee:"

    SetLocalizedCaption "Label23", "form.letter_creator.label.summary_letter_number", "Letter number:"

    SetLocalizedCaption "Label25", "form.letter_creator.label.summary_date", "Date:"

    SetLocalizedCaption "Label27", "form.letter_creator.label.summary_executor", "Executor:"

    SetLocalizedCaption "Label29", "form.letter_creator.label.summary_document_count", "Document count:"

    SetLocalizedCaption "Label30", "form.letter_creator.label.summary_attachments", "Attachments:"

    SetLocalizedCaption "Label31", "form.letter_creator.label.document_sum", "Document amount"

    SetLocalizedCaption "lblSelectedDocument", "form.letter_creator.label.selected_document", "Selected document:"

    SetLocalizedCaption "Frame1", "form.letter_creator.frame.address_details", "Address details"

    SetLocalizedCaption "Frame5", "form.letter_creator.frame.letter_summary", "Letter summary"

    SetLocalizedCaption "btnSaveNewAddress", "form.letter_creator.caption.save_address", "Save address"

    SetLocalizedCaption "btnClearSearch", "form.letter_creator.caption.clear_search", "Clear"

    SetLocalizedCaption "btnPrevious", "form.letter_creator.caption.back", "< Back"

    SetLocalizedCaption "btnNext", "form.letter_creator.caption.next", "Next >"

    SetLocalizedCaption "btnCancel", "form.letter_creator.caption.cancel", "Cancel"

    SetLocalizedCaption "btnEditAddress", "form.letter_creator.caption.edit_address", "Edit address"

    SetLocalizedCaption "btnDeleteAddress", "form.letter_creator.caption.delete_address", "Delete address"

    SetLocalizedCaption "btnLetterHistory", "form.letter_creator.caption.letter_history", "Letters History"



    mpgWizard.Pages(0).Caption = t("form.letter_creator.page.step_1", "Step 1: Addressee")

    mpgWizard.Pages(1).Caption = t("form.letter_creator.page.step_2", "Step 2: Letter")

    mpgWizard.Pages(2).Caption = t("form.letter_creator.page.step_3", "Step 3: Attachments")

    mpgWizard.Pages(3).Caption = t("form.letter_creator.page.step_4", "Step 4: Create")



    On Error GoTo 0

End Sub



Private Sub SetLocalizedCaption(controlName As String, localizationKey As String, fallbackText As String)

    SetControlCaption controlName, t(localizationKey, fallbackText)

End Sub



Private Sub SetControlCaption(controlName As String, captionText As String)

    SetResolvedControlCaption controlName, captionText

End Sub



Private Sub SetResolvedControlCaption(controlName As String, captionText As String)

    On Error Resume Next



    Dim ctrl As Object

    Set ctrl = ResolveNamedControl(controlName)

    If Not ctrl Is Nothing Then

        ctrl.Caption = captionText

    End If



    On Error GoTo 0

End Sub



Private Sub SetResolvedControlTip(controlName As String, tipText As String)

    On Error Resume Next



    Dim ctrl As Object

    Set ctrl = ResolveNamedControl(controlName)

    If Not ctrl Is Nothing Then

        ctrl.ControlTipText = tipText

    End If



    On Error GoTo 0

End Sub



Private Function ResolveNamedControl(controlName As String) As Object

    On Error Resume Next

    Set ResolveNamedControl = Me.Controls(controlName)

    On Error GoTo 0



    If Not ResolveNamedControl Is Nothing Then Exit Function



    Dim hostControl As control

    For Each hostControl In Me.Controls

        On Error Resume Next

        Set ResolveNamedControl = hostControl.Controls(controlName)

        On Error GoTo 0

        If Not ResolveNamedControl Is Nothing Then Exit Function

    Next hostControl

End Function



Private Sub EnsureAddressGroupControls()

    On Error GoTo EnsureError



    Dim addressFrame As Object

    Set addressFrame = ResolveNamedControl("Frame1")

    If addressFrame Is Nothing Then Exit Sub



    addressFrame.Height = 324



    Dim groupLabel As Object

    Set groupLabel = Nothing

    On Error Resume Next

    Set groupLabel = addressFrame.Controls(ADDRESS_GROUP_LABEL_NAME)

    On Error GoTo EnsureError



    If groupLabel Is Nothing Then

        Set groupLabel = addressFrame.Controls.Add("Forms.Label.1", ADDRESS_GROUP_LABEL_NAME, True)

    End If



    With groupLabel

        .Caption = GetAddressGroupLabelText()

        .Left = 30

        .Top = 270

        .Width = 84

        .Height = 12

        .BackStyle = 0

        .Font.Name = "Segoe UI"

        .Font.Size = 9

    End With



    Dim groupTextBox As Object

    Set groupTextBox = Nothing

    On Error Resume Next

    Set groupTextBox = addressFrame.Controls(ADDRESS_GROUP_TEXTBOX_NAME)

    On Error GoTo EnsureError



    If groupTextBox Is Nothing Then

        Set groupTextBox = addressFrame.Controls.Add("Forms.TextBox.1", ADDRESS_GROUP_TEXTBOX_NAME, True)

    End If



    With groupTextBox

        .Left = 126

        .Top = 264

        .Width = 276

        .Height = 24

        .backColor = RGB(255, 255, 255)

        .ControlTipText = GetAddressGroupTooltipText()

        .Font.Name = "Segoe UI"

        .Font.Size = 10

        .MultiLine = False

    End With



    If Not btnSaveNewAddress Is Nothing Then btnSaveNewAddress.Top = 372

    If Not btnEditAddress Is Nothing Then btnEditAddress.Top = 372

    If Not btnDeleteAddress Is Nothing Then btnDeleteAddress.Top = 372

    If Not mpgWizard Is Nothing Then mpgWizard.Height = 414

    On Error Resume Next

    If Not mpgWizard Is Nothing Then mpgWizard.Pages(0).Height = 390

    On Error GoTo EnsureError

    If Not btnPrevious Is Nothing Then btnPrevious.Top = 468

    If Not btnNext Is Nothing Then btnNext.Top = 468

    If Not btnLetterHistory Is Nothing Then btnLetterHistory.Top = 516

    If Not btnCancel Is Nothing Then btnCancel.Top = 516

    Me.Height = 600

    Exit Sub



EnsureError:

    Debug.Print "Address group controls setup error: " & Err.description

End Sub



Private Function GetAddressGroupLabelText() As String

    GetAddressGroupLabelText = BuildUnicodeText(1043, 1088, 1091, 1087, 1087, 1072, 32, 1072, 1076, 1088, 1077, 1089, 1072)

End Function



Private Function GetAddressGroupTooltipText() As String

    GetAddressGroupTooltipText = BuildUnicodeText(1054, 1073, 1097, 1072, 1103, 32, 1075, 1088, 1091, 1087, 1087, 1072, 32, 1076, 1083, 1103, 32, 1072, 1076, 1088, 1077, 1089, 1086, 1074, 32, 1089, 32, 1086, 1076, 1085, 1080, 1084, 32, 1087, 1086, 1095, 1090, 1086, 1074, 1099, 1084, 32, 1072, 1076, 1088, 1077, 1089, 1086, 1084, 46, 32, 1053, 1072, 1087, 1088, 1080, 1084, 1077, 1088, 58, 32, 53, 32, 1060, 1069, 1054)

End Function



Private Function BuildUnicodeText(ParamArray codePoints() As Variant) As String

    Dim i As Long



    BuildUnicodeText = ""

    For i = LBound(codePoints) To UBound(codePoints)

        BuildUnicodeText = BuildUnicodeText & ChrW(CLng(codePoints(i)))

    Next i

End Function



Private Sub txtAddresseePhone_Change()

    On Error Resume Next

    

    Dim currentValue As String, cursorPos As Long

    currentValue = Me.Controls("txtAddresseePhone").value

    cursorPos = Me.Controls("txtAddresseePhone").SelStart

    

    If Len(currentValue) >= 7 Then

        Dim formattedPhone As String

        formattedPhone = FormatPhoneNumber(currentValue)

        

        If formattedPhone <> currentValue Then

            Me.Controls("txtAddresseePhone").value = formattedPhone

            Me.Controls("txtAddresseePhone").SelStart = WorksheetFunction.Min(cursorPos + (Len(formattedPhone) - Len(currentValue)), Len(formattedPhone))

        End If

    End If

    

    On Error GoTo 0

End Sub



'------------------------------------------------------------

'  EXECUTORS

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

'  CLEAR DOCUMENT FIELDS

'------------------------------------------------------------

Private Sub ClearDocumentFields()

    On Error Resume Next

    If Not txtDocNumber Is Nothing Then txtDocNumber.value = ""

    If Not txtDocDate Is Nothing Then txtDocDate.value = ""

    If Not txtDocCopies Is Nothing Then txtDocCopies.value = ""

    If Not txtDocSheets Is Nothing Then txtDocSheets.value = ""

    If Not txtDocumentSum Is Nothing Then txtDocumentSum.value = ""

    On Error GoTo 0

End Sub



'=====================================================================

'                       NAVIGATION

'=====================================================================

Private Sub btnPrevious_Click()

    If mpgWizard.value > 0 Then SwitchToPage mpgWizard.value - 1

End Sub



Private Sub btnNext_Click()

    Dim cur As Integer: cur = mpgWizard.value

    

    If Not ValidatePage(cur) Then Exit Sub

    

    If cur = TOTAL_PAGES - 1 Then

        If ValidateForm Then

            UpdateSummaryInfo

            CreateWordLetter

            SaveLetterToDatabase

            MsgBox t("form.letter_creator.msg.letter_created", "Letter created successfully!"), vbInformation

            Unload Me

        End If

    Else

        If cur = 2 Then UpdateSummaryInfo

        SwitchToPage cur + 1

    End If

End Sub



Private Sub SwitchToPage(pg As Integer)

    If pg < 0 Or pg > TOTAL_PAGES - 1 Then Exit Sub

    

    mpgWizard.value = pg

    lblProgressInfo.Caption = BuildCreatorProgressCaption(pg + 1, TOTAL_PAGES)

    

    btnPrevious.Enabled = (pg > 0)

    

    If pg = TOTAL_PAGES - 1 Then

        btnNext.Caption = t("form.letter_creator.caption.create_letter", "CREATE LETTER")

        btnNext.backColor = RGB(76, 175, 80)

        btnNext.ForeColor = RGB(255, 255, 255)

        btnNext.Font.Bold = True

        btnNext.Font.Size = 11

    Else

        btnNext.Caption = t("form.letter_creator.caption.next", "Next >")

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

'                STEPS VALIDATION

'=====================================================================

Private Function ValidatePage(pg As Integer) As Boolean

    ValidatePage = False



    Dim focusControlName As String

    Dim validationMessage As String



    validationMessage = ValidateCreatorPage( _
        pg, _
        GetControlText("txtAddressee"), _
        GetControlText("txtCity"), _
        GetControlText("txtRegion"), _
        GetControlText("txtPostalCode"), _
        GetControlText("txtAddresseePhone"), _
        GetControlText("txtLetterNumber"), _
        GetControlText("txtLetterDate"), _
        GetControlText("cmbExecutor"), _
        documentsList.count, _
        focusControlName)



    If Len(validationMessage) > 0 Then

        ShowValidationFailure validationMessage, focusControlName

        Exit Function

    End If



    If pg = 0 Then

        ValidateAndUpdateSelectedAddress

    End If

    

    ValidatePage = True

End Function



'=====================================================================

'                PREFIX PROTECTION IN LETTER NUMBER

'=====================================================================

Private Sub txtLetterNumber_Change()

    On Error Resume Next

    If Not txtLetterNumber Is Nothing Then

        Dim currentValue As String

        currentValue = txtLetterNumber.value

        

        If Left(currentValue, 2) <> "7/" Then

            Dim numericPart As String

            numericPart = Replace(currentValue, "7/", "")

            

            txtLetterNumber.value = "7/" & numericPart

            txtLetterNumber.SelStart = Len(txtLetterNumber.value)

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

'      STEP 1 - addressee search and selection

'=====================================================================

Private Sub txtAddressSearch_Change()

    On Error Resume Next

    

    If Not Me.Controls("lstAddresses") Is Nothing Then

        Me.Controls("lstAddresses").Clear

        Set currentAddressSearchResults = New Collection

        

        ResetAddressFormState

        

        If Len(Trim(Me.Controls("txtAddressSearch").value)) > 0 Then

            Dim res As Collection, i As Long

            Set res = SearchAddresses(Me.Controls("txtAddressSearch").value)

            Set currentAddressSearchResults = res

            For i = 1 To res.count

                Me.Controls("lstAddresses").AddItem GetAddressSearchResultDisplayText(res(i))

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

    

    Debug.Print "Address form state reset"

    

    On Error GoTo 0

End Sub



Private Sub lstAddresses_Click()

    On Error GoTo SelectError

    

    If lstAddresses Is Nothing Or lstAddresses.listIndex < 0 Then Exit Sub

    If currentAddressSearchResults Is Nothing Then Exit Sub

    If lstAddresses.listIndex + 1 > currentAddressSearchResults.count Then Exit Sub

    

    Dim searchResult As Variant

    Dim addressArray As Variant

    Dim rowNumber As Long

    Dim errorMessage As String

    searchResult = currentAddressSearchResults(lstAddresses.listIndex + 1)

    

    If Not TryGetAddressSearchSelection(searchResult, addressArray, rowNumber, errorMessage) Then

        MsgBox errorMessage, vbExclamation

        Exit Sub

    End If



    Dim freshAddressArray As Variant

    If TryLoadAddressRowByNumber(rowNumber, freshAddressArray) Then

        addressArray = freshAddressArray

    End If

    

    LoadAddressForEditing addressArray, rowNumber

    

    Exit Sub



SelectError:

    MsgBox t("form.letter_creator.msg.address_select_error", "Error selecting address: ") & Err.description, vbExclamation

End Sub



Public Sub LoadAddressForEditing(addressArray As Variant, rowNumber As Long)

    ApplyAddressPartsToControls addressArray

    selectedAddressRow = rowNumber



    If Not btnSaveNewAddress Is Nothing Then btnSaveNewAddress.Enabled = False

    If Not btnEditAddress Is Nothing Then btnEditAddress.Enabled = True

    If Not btnDeleteAddress Is Nothing Then btnDeleteAddress.Enabled = True



    SafeSetFocus "txtAddressee"

End Sub



Private Sub btnSaveNewAddress_Click()

    On Error Resume Next



    Dim addressArray As Variant

    Dim validationMessage As String

    

    addressArray = CreateAddressArray()

    validationMessage = ValidateAddressCreateRequest(GetControlText("txtAddressee"), IsAddressDuplicate(addressArray))

    If Len(validationMessage) > 0 Then

        MsgBox validationMessage, vbExclamation

        Exit Sub

    End If

    

    SaveNewAddress addressArray

    MsgBox t("form.letter_creator.msg.address_saved", "Address saved."), vbInformation

    

    ClearAddressCache

    

    On Error GoTo 0

End Sub



'=====================================================================

'      ADDRESS EDITING BUTTONS

'=====================================================================

Private Sub btnEditAddress_Click()

    On Error Resume Next

    

    ValidateAndUpdateSelectedAddress

    

    Dim addressArray As Variant

    Dim validationMessage As String

    addressArray = CreateAddressArray()

    

    validationMessage = ValidateAddressEditRequest(selectedAddressRow, IsAddressDuplicate(addressArray, selectedAddressRow))

    If Len(validationMessage) > 0 Then

        MsgBox validationMessage, vbExclamation

        Exit Sub

    End If

    

    UpdateExistingAddress selectedAddressRow, addressArray

    

    ClearAddressCache

    txtAddressSearch_Change

    

    MsgBox t("form.letter_creator.msg.address_updated", "Address updated successfully."), vbInformation

    On Error GoTo 0

End Sub



Private Sub btnDeleteAddress_Click()

    On Error GoTo DeleteError

    

    Dim validationMessage As String

    validationMessage = ValidateAddressDeleteRequest(selectedAddressRow)

    

    If Len(validationMessage) > 0 Then

        MsgBox validationMessage, vbExclamation

        Exit Sub

    End If

    

    If MsgBox(t("form.letter_creator.msg.address_delete_confirm", "Are you sure you want to delete this address?"), vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then

        DeleteExistingAddress selectedAddressRow

        MsgBox t("form.letter_creator.msg.address_deleted", "Address deleted successfully."), vbInformation

        

        ClearAddressFields

        ClearAddressCache

        

        selectedAddressRow = 0

        btnEditAddress.Enabled = False

        btnDeleteAddress.Enabled = False

    End If

    

    Exit Sub

    

DeleteError:

    MsgBox t("form.letter_creator.msg.address_delete_error", "Error deleting address: ") & Err.description, vbCritical

End Sub



Private Sub ClearAddressFields()

    On Error Resume Next

    If Not txtAddressee Is Nothing Then txtAddressee.value = ""

    If Not txtStreet Is Nothing Then txtStreet.value = ""

    If Not txtCity Is Nothing Then txtCity.value = ""

    If Not txtDistrict Is Nothing Then txtDistrict.value = ""

    If Not txtRegion Is Nothing Then txtRegion.value = ""

    If Not txtPostalCode Is Nothing Then txtPostalCode.value = ""

    If Not txtAddresseePhone Is Nothing Then txtAddresseePhone.value = ""

    SetControlValue ADDRESS_GROUP_TEXTBOX_NAME, ""

    On Error GoTo 0

End Sub



'=====================================================================

'      STEP 3 - adding attachments

'=====================================================================

Private Sub txtAttachmentSearch_Change()

    On Error Resume Next

    If Not lstAvailableAttachments Is Nothing Then

        lstAvailableAttachments.Clear

        If Len(Trim(txtAttachmentSearch.value)) > 0 Then

            Dim res As Collection, i As Long

            Set res = GetCachedAttachments(txtAttachmentSearch.value)

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

    If lstAvailableAttachments.listIndex >= 0 And Not txtDocNumber Is Nothing Then

        txtDocNumber.SetFocus

    End If

    On Error GoTo 0

End Sub



'=====================================================================

'      ADDING ATTACHMENTS WITH SUM

'=====================================================================

Private Sub btnAddAttachment_Click()

    On Error Resume Next

    

    If lstAvailableAttachments Is Nothing Or lstAvailableAttachments.listIndex < 0 Then

        MsgBox t("form.letter_creator.msg.select_document_left", "Select a document in the left list."), vbExclamation

        Exit Sub

    End If

    

    Dim docArr As Variant

    docArr = CreateDocumentArrayWithSum( _
        lstAvailableAttachments.List(lstAvailableAttachments.listIndex), _
        Trim(IIf(txtDocNumber Is Nothing, "", txtDocNumber.value)), _
        Trim(IIf(txtDocDate Is Nothing, "", txtDocDate.value)), _
        Trim(IIf(txtDocCopies Is Nothing, "", txtDocCopies.value)), _
        Trim(IIf(txtDocSheets Is Nothing, "", txtDocSheets.value)), _
        Trim(IIf(txtDocumentSum Is Nothing, "", txtDocumentSum.value)))

    

    documentsList.Add docArr

    SyncSelectedAttachmentsList

    

    ClearDocumentFields

    On Error GoTo 0

End Sub



Private Sub btnRemoveAttachment_Click()

    On Error Resume Next

    

    If lstSelectedAttachments Is Nothing Or lstSelectedAttachments.listIndex < 0 Then

        MsgBox t("form.letter_creator.msg.select_document_right", "Select a document in the right list."), vbExclamation

        Exit Sub

    End If

    

    Dim selectedIndex As Integer

    selectedIndex = lstSelectedAttachments.listIndex

    

    If selectedIndex + 1 <= documentsList.count Then

        documentsList.Remove selectedIndex + 1

        SyncSelectedAttachmentsList

    End If

    

    On Error GoTo 0

End Sub



'=====================================================================

'                   CONTEXT MENU

'=====================================================================

Private Sub lstSelectedAttachments_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Button = 2 And lstSelectedAttachments.listIndex >= 0 Then

        contextMenuSelectedIndex = lstSelectedAttachments.listIndex

        ShowSimpleContextMenu

    End If

End Sub



Private Sub ShowSimpleContextMenu()

    On Error GoTo MenuError

    

    Dim menuChoice As String

    menuChoice = InputBox(GetDocumentActionsMenuPrompt(), GetDocumentActionsMenuTitle(), "1")

    

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

            If UBound(docArray) >= DocumentIndexSheets Then

                txtDocNumber.value = docArray(DocumentIndexNumber)

                txtDocDate.value = docArray(DocumentIndexDate)

                txtDocCopies.value = docArray(DocumentIndexCopies)

                txtDocSheets.value = docArray(DocumentIndexSheets)

            End If

            

            If UBound(docArray) >= DocumentIndexSum And Not txtDocumentSum Is Nothing Then

                txtDocumentSum.value = docArray(DocumentIndexSum)

            End If

        End If

    End If

    Exit Sub

    

EditError:

End Sub



Public Sub DuplicateDocument()

    On Error GoTo DuplicateError

    

    If contextMenuSelectedIndex >= 0 And contextMenuSelectedIndex < documentsList.count Then

        Dim sourceItem As Variant

        sourceItem = documentsList.item(contextMenuSelectedIndex + 1)

        

        Dim duplicateDoc As Variant

        duplicateDoc = DuplicateDocumentArray(sourceItem)

        

        documentsList.Add duplicateDoc

        SyncSelectedAttachmentsList

    End If

    Exit Sub

    

DuplicateError:

    MsgBox t("form.letter_creator.msg.duplicate_document_error", "Error duplicating document: ") & Err.description, vbCritical

End Sub



Public Sub RemoveSelectedDocument()

    On Error GoTo RemoveError

    

    If contextMenuSelectedIndex >= 0 And contextMenuSelectedIndex < documentsList.count Then

        documentsList.Remove contextMenuSelectedIndex + 1

        SyncSelectedAttachmentsList

    End If

    Exit Sub

    

RemoveError:

End Sub



Public Sub MoveDocumentUp()

    On Error GoTo MoveUpError

    

    If contextMenuSelectedIndex > 0 Then

        MoveDocumentCollectionItemUp documentsList, contextMenuSelectedIndex + 1

        RefreshDocumentsList

        lstSelectedAttachments.listIndex = contextMenuSelectedIndex - 1

    End If

    Exit Sub

    

MoveUpError:

End Sub



Public Sub MoveDocumentDown()

    On Error GoTo MoveDownError

    

    If contextMenuSelectedIndex < documentsList.count - 1 Then

        MoveDocumentCollectionItemDown documentsList, contextMenuSelectedIndex + 1

        RefreshDocumentsList

        lstSelectedAttachments.listIndex = contextMenuSelectedIndex + 1

    End If

    Exit Sub

    

MoveDownError:

End Sub



Private Sub RefreshDocumentsList()

    SyncSelectedAttachmentsList

End Sub



Private Sub SyncSelectedAttachmentsList()

    lstSelectedAttachments.Clear

    

    Dim displayItems As Collection

    Dim i As Long

    

    Set displayItems = GetDocumentDisplayItems(documentsList)

    For i = 1 To displayItems.count

        lstSelectedAttachments.AddItem displayItems(i)

    Next i

    

    UpdateSelectedDocumentsCaption

End Sub



Private Sub UpdateSelectedDocumentsCaption()

    If Not lblAttachmentsCount Is Nothing Then

        lblAttachmentsCount.Caption = BuildCreatorSelectedDocumentsCaption(documentsList.count)

    End If

End Sub



'=====================================================================

'      STEP 4 - summary and letter creation

'=====================================================================

Private Sub UpdateSummaryInfo()

    On Error Resume Next

    

    If Not lblSummaryRecipient Is Nothing Then

        lblSummaryRecipient.Caption = IIf(txtAddressee Is Nothing, "", txtAddressee.value)

    End If

    

    If Not lblSummaryNumber Is Nothing Then

        lblSummaryNumber.Caption = IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.value)

    End If

    

    If Not lblSummaryDate Is Nothing Then

        lblSummaryDate.Caption = IIf(txtLetterDate Is Nothing, "", txtLetterDate.value)

    End If

    

    If Not lblSummaryExecutor Is Nothing Then

        lblSummaryExecutor.Caption = IIf(cmbExecutor Is Nothing, "", cmbExecutor.value)

    End If

    

    If Not lblSummaryDocsCount Is Nothing Then

        lblSummaryDocsCount.Caption = CStr(documentsList.count)

    End If

    

    If Not txtFinalAttachments Is Nothing Then

        Dim attachmentText As String

        attachmentText = BuildSummaryAttachmentsText(documentsList)

        txtFinalAttachments.value = attachmentText

    End If

    

    On Error GoTo 0

End Sub



'=====================================================================

'  GLOBAL VALIDATION BEFORE CREATION

'=====================================================================

Private Function ValidateForm() As Boolean

    ValidateForm = False



    Dim focusControlName As String

    Dim validationMessage As String



    validationMessage = ValidateCreatorSubmission( _
        GetControlText("txtAddressee"), _
        GetControlText("txtCity"), _
        GetControlText("txtRegion"), _
        GetControlText("txtPostalCode"), _
        GetControlText("txtLetterNumber"), _
        GetControlText("txtLetterDate"), _
        GetControlText("cmbExecutor"), _
        documentsList.count, _
        focusControlName)



    If Len(validationMessage) > 0 Then

        SwitchToPage GetPageIndexForControl(focusControlName)

        ShowValidationFailure validationMessage, focusControlName

        Exit Function

    End If

    

    ValidateForm = True

End Function



'=====================================================================

'  AUXILIARY FUNCTIONS FOR ADDRESS

'=====================================================================

Private Function CreateAddressArray() As Variant

    Dim arr(AddressIndexGroup) As String

    

    arr(AddressIndexAddressee) = GetControlText("txtAddressee")

    arr(AddressIndexStreet) = GetControlText("txtStreet")

    arr(AddressIndexCity) = GetControlText("txtCity")

    arr(AddressIndexDistrict) = GetControlText("txtDistrict")

    arr(AddressIndexRegion) = GetControlText("txtRegion")

    arr(AddressIndexPostalCode) = GetControlText("txtPostalCode")

    arr(AddressIndexPhone) = GetControlText("txtAddresseePhone")

    arr(AddressIndexGroup) = GetControlText(ADDRESS_GROUP_TEXTBOX_NAME)

    

    CreateAddressArray = arr

End Function



Private Sub ApplyAddressPartsToControls(addressParts As Variant)

    SetControlValue "txtAddressee", CStr(addressParts(AddressIndexAddressee))

    SetControlValue "txtStreet", CStr(addressParts(AddressIndexStreet))

    SetControlValue "txtCity", CStr(addressParts(AddressIndexCity))

    SetControlValue "txtDistrict", CStr(addressParts(AddressIndexDistrict))

    SetControlValue "txtRegion", CStr(addressParts(AddressIndexRegion))

    SetControlValue "txtPostalCode", CStr(addressParts(AddressIndexPostalCode))

    SetControlValue "txtAddresseePhone", CStr(addressParts(AddressIndexPhone))

    SetControlValue ADDRESS_GROUP_TEXTBOX_NAME, CStr(addressParts(AddressIndexGroup))

End Sub



Private Function GetControlText(controlName As String) As String

    On Error Resume Next

    Dim ctrl As Object

    Set ctrl = ResolveNamedControl(controlName)

    If Not ctrl Is Nothing Then

        GetControlText = Trim(CStr(ctrl.value))

    Else

        GetControlText = ""

    End If

    On Error GoTo 0

End Function



Private Sub SetControlValue(controlName As String, controlValue As String)

    On Error Resume Next

    Dim ctrl As Object

    Set ctrl = ResolveNamedControl(controlName)

    If Not ctrl Is Nothing Then

        ctrl.value = controlValue

    End If

    On Error GoTo 0

End Sub



Private Sub ShowValidationFailure(messageText As String, focusControlName As String)

    MsgBox messageText, vbExclamation

    SafeSetFocus focusControlName

End Sub



Private Function GetPageIndexForControl(controlName As String) As Integer

    Select Case controlName

        Case "txtAddressee", "txtCity", "txtRegion", "txtPostalCode", "txtAddresseePhone", ADDRESS_GROUP_TEXTBOX_NAME

            GetPageIndexForControl = 0

        Case "txtLetterNumber", "txtLetterDate", "cmbExecutor"

            GetPageIndexForControl = 1

        Case Else

            GetPageIndexForControl = 2

    End Select

End Function



'=====================================================================

'      CREATING LETTER IN WORD

'=====================================================================

Private Sub CreateWordLetter()

    On Error GoTo ErrorHandler

    

    CreateLetterDocument _
        IIf(txtAddressee Is Nothing, "", txtAddressee.value), _
        CreateAddressArray(), _
        IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.value), _
        IIf(txtLetterDate Is Nothing, "", txtLetterDate.value), _
        IIf(cmbExecutor Is Nothing, "", cmbExecutor.value), _
        IIf(cmbDocumentType Is Nothing, "", ResolveDocumentTypeStorageValue(cmbDocumentType.value)), _
        IsAlternateLetterTypeSelection(IIf(cmbLetterType Is Nothing, "", cmbLetterType.value)), _
        documentsList

    Exit Sub



ErrorHandler:

    MsgBox t("form.letter_creator.msg.create_letter_error", "Error creating letter: ") & Err.description, vbCritical

End Sub



'=====================================================================

'      SAVING TO DATABASE WITH SUM

'=====================================================================

Private Sub SaveLetterToDatabase()

    SaveLetterInfoWithSum IIf(txtAddressee Is Nothing, "", txtAddressee.value), _
                          IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.value), _
                          ResolveLetterDateOrToday(GetControlText("txtLetterDate")), documentsList, _
                          IIf(cmbExecutor Is Nothing, "", cmbExecutor.value), _
                          IIf(cmbDocumentType Is Nothing, "", _
                              IIf(cmbDocumentType.listIndex >= 0, ResolveDocumentTypeStorageValue(cmbDocumentType.value), ""))

End Sub



'=====================================================================

'      AUTO-UPDATE ADDRESSES

'=====================================================================

Private Sub AutoUpdateAddressIfChanged()

    On Error Resume Next

    

    If ShouldSkipAddressAutoUpdate() Then Exit Sub

    If selectedAddressRow <= 1 Then Exit Sub

    

    Dim currentAddress As Variant

    currentAddress = CreateAddressArray()

    

    If HasAddressDataChanged(selectedAddressRow, currentAddress) Then

        UpdateExistingAddress selectedAddressRow, currentAddress

        Debug.Print "Address automatically updated in row " & selectedAddressRow

        ClearAddressCache

    End If

    

    On Error GoTo 0

End Sub



Private Sub ValidateAndUpdateSelectedAddress()

    On Error Resume Next

    

    If selectedAddressRow > 1 Then

        If IsAddressReadyForAutoUpdate(GetControlText("txtCity"), GetControlText("txtRegion"), GetControlText("txtPostalCode")) Then

            AutoUpdateAddressIfChanged

        End If

    End If

    

    On Error GoTo 0

End Sub



'=====================================================================

'      MULTILINE FIELDS CONFIGURATION

'=====================================================================

Private Sub ConfigureMultilineTextBoxes()

    On Error Resume Next

    

    Dim ctrl As control

    Dim textboxNames As Variant

    Dim i As Long

    

    textboxNames = Array("txtAddressee", "txtStreet", "txtCity", "txtDistrict", "txtRegion", "txtPostalCode")

    

    For i = LBound(textboxNames) To UBound(textboxNames)

        Set ctrl = ResolveNamedControl(CStr(textboxNames(i)))

        If Not ctrl Is Nothing Then

            ctrl.MultiLine = True

            ctrl.WordWrap = True

            ctrl.ScrollBars = 2

            

            If ctrl.Height < 40 Then ctrl.Height = 35

        End If

    Next i

    

    On Error GoTo 0

End Sub



Private Sub AutoResizeTextBoxHeight(controlName As String)

    On Error Resume Next

    

    Dim ctrl As control

    Set ctrl = ResolveNamedControl(controlName)

    

    If Not ctrl Is Nothing Then

        Dim textLength As Long

        Dim linesCount As Long

        

        textLength = Len(ctrl.value)

        linesCount = Int(textLength / 40) + 1

        

        If linesCount < 1 Then linesCount = 1

        If linesCount > 4 Then linesCount = 4

        

        ctrl.Height = linesCount * 18 + 10

    End If

    

    On Error GoTo 0

End Sub



'=====================================================================

'      SAFE FOCUS SETTING

'=====================================================================

Private Sub SafeSetFocus(controlName As String)

    On Error Resume Next

    Dim ctrl As Object

    Set ctrl = ResolveNamedControl(controlName)

    If Not ctrl Is Nothing Then

        If ctrl.Enabled And ctrl.Visible Then

            ctrl.SetFocus

        End If

    End If

    On Error GoTo 0

End Sub



'=====================================================================

'  CANCEL AND CLOSE BUTTONS

'=====================================================================

Private Sub btnCancel_Click()

    isClosingForm = True

    If MsgBox(t("dialog.cancel_letter_creation", "Cancel letter creation?"), vbYesNo + vbQuestion) = vbYes Then

        ClearCache

        Unload Me

    Else

        isClosingForm = False

    End If

End Sub



Private Sub btnCancel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    skipNextAddressAutoUpdate = True

End Sub



Private Sub btnPrevious_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    skipNextAddressAutoUpdate = True

End Sub



Private Sub btnNext_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    skipNextAddressAutoUpdate = True

End Sub



Private Sub btnLetterHistory_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    skipNextAddressAutoUpdate = True

End Sub



Private Sub btnClearSearch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    skipNextAddressAutoUpdate = True

End Sub



Private Function ShouldSkipAddressAutoUpdate() As Boolean

    If skipNextAddressAutoUpdate Then

        skipNextAddressAutoUpdate = False

        ShouldSkipAddressAutoUpdate = True

        Exit Function

    End If



    If isClosingForm Then

        ShouldSkipAddressAutoUpdate = True

        Exit Function

    End If



    Dim activeControlName As String

    activeControlName = GetActiveControlName()



    Select Case activeControlName

        Case "btnCancel", "btnPrevious", "btnNext", "btnLetterHistory"

            ShouldSkipAddressAutoUpdate = True

        Case Else

            ShouldSkipAddressAutoUpdate = False

    End Select

End Function



Private Function GetActiveControlName() As String

    On Error Resume Next

    GetActiveControlName = CStr(Me.ActiveControl.Name)

    On Error GoTo 0

End Function



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    isClosingForm = True

    If documentsList.count > 0 Then

        If MsgBox(t("dialog.discard_unsaved_documents", "Unsaved documents will be lost. Close anyway?"), vbYesNo + vbQuestion) = vbNo Then

            isClosingForm = False

            Cancel = True

        Else

            ClearCache

        End If

    Else

        ClearCache

    End If

End Sub




