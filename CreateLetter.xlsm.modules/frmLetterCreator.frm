VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLetterCreator 
   Caption         =   "Letter Builder v1.6.2"
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
' Form    : frmLetterCreator v1.6.2 - Thin-shell MultiPage wizard for letter creation
' Version : 1.6.2 - 26.03.2026
' Author  : Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose : UI orchestration for letter creation, address entry, attachments, and summary flow
' ======================================================================

Option Explicit

'------------------------------------------------------------
'  GLOBAL VARIABLES
'------------------------------------------------------------
Private Const TOTAL_PAGES As Integer = 4
Public selectedAddressRow As Long
Private documentsList As Collection
Private contextMenuSelectedIndex As Integer

'------------------------------------------------------------
'  FORM INITIALIZATION
'------------------------------------------------------------
Private Sub UserForm_Initialize()
    Set documentsList = New Collection
    contextMenuSelectedIndex = -1
    
    ConfigureFormAppearance
    ApplyEnglishStaticCaptions
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
            .ControlTipText = "Document sum in rubles (optional). For example: 125000"
            .Value = ""
            .backColor = RGB(255, 255, 255)
        End With
        Debug.Print "Document sum field configured"
    End If
    
    On Error GoTo 0
End Sub

Private Sub InitializeProgressInfo()
    lblProgressInfo.Caption = "Step 1 of " & TOTAL_PAGES
    lblAttachmentsCount.Caption = "Selected documents: 0"
End Sub

'------------------------------------------------------------
'  CONTROL VALUES INITIALIZATION
'------------------------------------------------------------
Private Sub InitializeControlValues()
    On Error Resume Next
    
    Me.Controls("txtLetterDate").Value = Format(Date, "dd.mm.yyyy")
    Me.Controls("txtLetterNumber").Value = "7/"
    
    With Me.Controls("cmbDocumentType")
        .Clear
        .AddItem "Third-party confirmed documents"
        .AddItem "Own for confirmation"
        .ListIndex = 0
    End With
    
    With Me.Controls("cmbLetterType")
        .Clear
        .AddItem "Regular"
        .AddItem "FOU (For Official Use)"
        .ListIndex = 0
    End With
    
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
    MsgBox "Error opening history form: " & Err.description, vbCritical
End Sub

'------------------------------------------------------------
'  CLEAR SEARCH BUTTON
'------------------------------------------------------------
Private Sub btnClearSearch_Click()
    On Error Resume Next
    
    Me.Controls("txtAddressSearch").Value = ""
    Me.Controls("lstAddresses").Clear
    
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
    Dim ctrl As Control
    
    addressFields = Array("txtAddressee", "txtStreet", "txtCity", "txtDistrict", "txtRegion", "txtPostalCode", "txtAddresseePhone")
    
    For i = LBound(addressFields) To UBound(addressFields)
        Set ctrl = Me.Controls(addressFields(i))
        If Not ctrl Is Nothing Then
            ctrl.Value = ""
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
'  APPEARANCE CONFIGURATION
'------------------------------------------------------------
Private Sub ConfigureFormAppearance()
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 10
    Me.Caption = "Letter Builder v1.6.2"
    
    On Error Resume Next
    
    If Not lstSelectedAttachments Is Nothing Then
        lstSelectedAttachments.Font.Size = 9
        lstSelectedAttachments.ControlTipText = "Hover over the item to see the full name"
        lstSelectedAttachments.IntegralHeight = False
    End If
    
    If Not btnEditAddress Is Nothing Then
        btnEditAddress.Caption = "Edit address"
        btnEditAddress.ControlTipText = "Edit selected address"
        btnEditAddress.Enabled = False
    End If
    
    If Not btnDeleteAddress Is Nothing Then
        btnDeleteAddress.Caption = "Delete address"
        btnDeleteAddress.ControlTipText = "Delete selected address"
        btnDeleteAddress.Enabled = False
    End If
    
    If Not txtAddresseePhone Is Nothing Then
        txtAddresseePhone.ControlTipText = "Addressee phone (format: 8-xxx-xxx-xx-xx)"
        txtAddresseePhone.Enabled = True
        txtAddresseePhone.backColor = RGB(255, 255, 255)
    End If
    
    If Not btnLetterHistory Is Nothing Then
        With btnLetterHistory
            .Caption = "Letters History"
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .backColor = RGB(156, 39, 176)
            .ForeColor = RGB(255, 255, 255)
            .ControlTipText = "Open sent letters history form"
        End With
    End If
    
    On Error GoTo 0
    
    txtAddressSearch.ControlTipText = "Enter part of the name to search for the addressee"
    txtLetterNumber.ControlTipText = "Enter the number after 7/ (for example: 125 becomes 7/125)"
    txtLetterDate.ControlTipText = "Format: dd.mm.yyyy"
End Sub

Private Sub ApplyEnglishStaticCaptions()
    On Error Resume Next

    SetControlCaption "lblStep1", "Stage:"
    SetControlCaption "lblStep2", ""
    SetControlCaption "lblStep3", ""
    SetControlCaption "lblStep4", ""
    SetControlCaption "lblStep5", ""
    SetControlCaption "lblCurrentAction", "Current action"
    SetControlCaption "Label1", "Search existing addressee"
    SetControlCaption "Label2", "City"
    SetControlCaption "Label3", "District"
    SetControlCaption "Label4", "Region"
    SetControlCaption "Label5", "Postal code"
    SetControlCaption "Label6", "Executor"
    SetControlCaption "Label7", "Letter date"
    SetControlCaption "Label8", "Letter number"
    SetControlCaption "Label9", "Search attachment"
    SetControlCaption "Label10", "Selected attachments"
    SetControlCaption "Label11", "Document ownership"
    SetControlCaption "Label13", "Date"
    SetControlCaption "Label14", "Copies"
    SetControlCaption "Label15", "Sheets"
    SetControlCaption "Label16", "Found addresses"
    SetControlCaption "Label17", "Street, house"
    SetControlCaption "Label18", "Addressee"
    SetControlCaption "Label19", "Available attachments"
    SetControlCaption "Label20", "Number"
    SetControlCaption "Label21", "Addressee:"
    SetControlCaption "Label23", "Letter number:"
    SetControlCaption "Label25", "Date:"
    SetControlCaption "Label27", "Executor:"
    SetControlCaption "Label29", "Document count:"
    SetControlCaption "Label30", "Attachments:"
    SetControlCaption "Label31", "Document sum"
    SetControlCaption "lblSelectedDocument", "Selected document:"
    SetControlCaption "Frame1", "Address details"
    SetControlCaption "Frame5", "Letter summary"
    SetControlCaption "btnSaveNewAddress", "Save address"
    SetControlCaption "btnClearSearch", "Clear"
    SetControlCaption "btnPrevious", "< Back"
    SetControlCaption "btnNext", "Next >"
    SetControlCaption "btnCancel", "Cancel"
    SetControlCaption "btnEditAddress", "Edit address"
    SetControlCaption "btnDeleteAddress", "Delete address"
    SetControlCaption "btnLetterHistory", "Letters History"

    mpgWizard.Pages(0).Caption = "Step 1: Addressee"
    mpgWizard.Pages(1).Caption = "Step 2: Letter"
    mpgWizard.Pages(2).Caption = "Step 3: Attachments"
    mpgWizard.Pages(3).Caption = "Step 4: Create"

    On Error GoTo 0
End Sub

Private Sub SetControlCaption(controlName As String, captionText As String)
    On Error Resume Next

    Dim ctrl As Control
    Set ctrl = Me.Controls(controlName)
    If Not ctrl Is Nothing Then
        ctrl.Caption = captionText
    End If

    On Error GoTo 0
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
    If Not txtDocNumber Is Nothing Then txtDocNumber.Value = ""
    If Not txtDocDate Is Nothing Then txtDocDate.Value = ""
    If Not txtDocCopies Is Nothing Then txtDocCopies.Value = ""
    If Not txtDocSheets Is Nothing Then txtDocSheets.Value = ""
    If Not txtDocumentSum Is Nothing Then txtDocumentSum.Value = ""
    On Error GoTo 0
End Sub

'=====================================================================
'                       NAVIGATION
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
            MsgBox "Letter created successfully!", vbInformation
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
    lblProgressInfo.Caption = "Step " & pg + 1 & " of " & TOTAL_PAGES
    
    btnPrevious.Enabled = (pg > 0)
    
    If pg = TOTAL_PAGES - 1 Then
        btnNext.Caption = "CREATE LETTER"
        btnNext.backColor = RGB(76, 175, 80)
        btnNext.ForeColor = RGB(255, 255, 255)
        btnNext.Font.Bold = True
        btnNext.Font.Size = 11
    Else
        btnNext.Caption = "Next >"
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
'      STEP 1 - addressee search and selection
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
    
    Debug.Print "Address form state reset"
    
    On Error GoTo 0
End Sub

Private Sub lstAddresses_Click()
    On Error Resume Next
    
    If lstAddresses Is Nothing Or lstAddresses.ListIndex < 0 Then Exit Sub
    
    Dim itm As String, parts As Variant
    itm = lstAddresses.List(lstAddresses.ListIndex)
    
    If InStr(itm, " | ") = 0 Then
        MsgBox "Invalid address record format.", vbExclamation
        Exit Sub
    End If
    
    parts = Split(itm, " | ")
    If UBound(parts) < 7 Then
        MsgBox "Address data is incomplete.", vbExclamation
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
        MsgBox "Enter the addressee name.", vbExclamation
        Exit Sub
    End If
    
    If IsAddressDuplicate(CreateAddressArray) Then
        MsgBox "This address already exists.", vbExclamation
        Exit Sub
    End If
    
    SaveNewAddress CreateAddressArray
    MsgBox "Address saved.", vbInformation
    
    ClearAddressCache
    
    On Error GoTo 0
End Sub

'=====================================================================
'      ADDRESS EDITING BUTTONS
'=====================================================================
Private Sub btnEditAddress_Click()
    On Error Resume Next
    
    If selectedAddressRow <= 1 Then
        MsgBox "Select an address to edit.", vbExclamation
        Exit Sub
    End If
    
    ValidateAndUpdateSelectedAddress
    
    Dim addressArray As Variant
    addressArray = CreateAddressArray()
    
    If IsAddressDuplicate(addressArray, selectedAddressRow) Then
        MsgBox "An address with the same data already exists.", vbExclamation
        Exit Sub
    End If
    
    UpdateExistingAddress selectedAddressRow, addressArray
    
    ClearAddressCache
    txtAddressSearch_Change
    
    MsgBox "Address updated successfully.", vbInformation
    On Error GoTo 0
End Sub

Private Sub btnDeleteAddress_Click()
    On Error GoTo DeleteError
    
    If selectedAddressRow = 0 Then
        MsgBox "Select an address to delete.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this address?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        DeleteExistingAddress selectedAddressRow
        MsgBox "Address deleted successfully.", vbInformation
        
        ClearAddressFields
        ClearAddressCache
        
        selectedAddressRow = 0
        btnEditAddress.Enabled = False
        btnDeleteAddress.Enabled = False
    End If
    
    Exit Sub
    
DeleteError:
    MsgBox "Error deleting address: " & Err.description, vbCritical
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
'      STEP 3 - adding attachments
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
'      ADDING ATTACHMENTS WITH SUM
'=====================================================================
Private Sub btnAddAttachment_Click()
    On Error Resume Next
    
    If lstAvailableAttachments Is Nothing Or lstAvailableAttachments.ListIndex < 0 Then
        MsgBox "Select a document in the left list.", vbExclamation
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
    SyncSelectedAttachmentsList
    
    ClearDocumentFields
    On Error GoTo 0
End Sub

Private Sub btnRemoveAttachment_Click()
    On Error Resume Next
    
    If lstSelectedAttachments Is Nothing Or lstSelectedAttachments.ListIndex < 0 Then
        MsgBox "Select a document in the right list.", vbExclamation
        Exit Sub
    End If
    
    Dim selectedIndex As Integer
    selectedIndex = lstSelectedAttachments.ListIndex
    
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
    If Button = 2 And lstSelectedAttachments.ListIndex >= 0 Then
        contextMenuSelectedIndex = lstSelectedAttachments.ListIndex
        ShowSimpleContextMenu
    End If
End Sub

Private Sub ShowSimpleContextMenu()
    On Error GoTo MenuError
    
    Dim menuChoice As String
    menuChoice = InputBox("Select action:" & vbCrLf & _
                         "1 - Edit details" & vbCrLf & _
                         "2 - Duplicate document" & vbCrLf & _
                         "3 - Remove from list" & vbCrLf & _
                         "4 - Move up" & vbCrLf & _
                         "5 - Move down", _
                         "Document actions", "1")
    
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
        Dim sourceItem As Variant
        sourceItem = documentsList.item(contextMenuSelectedIndex + 1)
        
        Dim duplicateDoc As Variant
        duplicateDoc = DuplicateDocumentArray(sourceItem)
        
        documentsList.Add duplicateDoc
        SyncSelectedAttachmentsList
    End If
    Exit Sub
    
DuplicateError:
    MsgBox "Error duplicating document: " & Err.description, vbCritical
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
        lstSelectedAttachments.ListIndex = contextMenuSelectedIndex - 1
    End If
    Exit Sub
    
MoveUpError:
End Sub

Public Sub MoveDocumentDown()
    On Error GoTo MoveDownError
    
    If contextMenuSelectedIndex < documentsList.count - 1 Then
        MoveDocumentCollectionItemDown documentsList, contextMenuSelectedIndex + 1
        RefreshDocumentsList
        lstSelectedAttachments.ListIndex = contextMenuSelectedIndex + 1
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
        lblAttachmentsCount.Caption = "Selected documents: " & documentsList.count
    End If
End Sub

'=====================================================================
'      STEP 4 - summary and letter creation
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
        attachmentText = BuildSummaryAttachmentsText(documentsList)
        txtFinalAttachments.Value = attachmentText
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

Private Function GetControlText(controlName As String) As String
    On Error Resume Next
    GetControlText = Trim(CStr(Me.Controls(controlName).Value))
    On Error GoTo 0
End Function

Private Sub ShowValidationFailure(messageText As String, focusControlName As String)
    MsgBox messageText, vbExclamation
    SafeSetFocus focusControlName
End Sub

Private Function GetPageIndexForControl(controlName As String) As Integer
    Select Case controlName
        Case "txtAddressee", "txtCity", "txtRegion", "txtPostalCode", "txtAddresseePhone"
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
        IIf(txtAddressee Is Nothing, "", txtAddressee.Value), _
        CreateAddressArray(), _
        IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value), _
        IIf(txtLetterDate Is Nothing, "", txtLetterDate.Value), _
        IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value), _
        IIf(cmbDocumentType Is Nothing, "", cmbDocumentType.Value), _
        (Not cmbLetterType Is Nothing And cmbLetterType.ListIndex = 1), _
        documentsList
    Exit Sub

ErrorHandler:
    MsgBox "Error creating letter: " & Err.Description, vbCritical
End Sub

'=====================================================================
'      SAVING TO DATABASE WITH SUM
'=====================================================================
Private Sub SaveLetterToDatabase()
    SaveLetterInfoWithSum IIf(txtAddressee Is Nothing, "", txtAddressee.Value), _
                          IIf(txtLetterNumber Is Nothing, "", txtLetterNumber.Value), _
                          ResolveLetterDateOrToday(GetControlText("txtLetterDate")), documentsList, _
                          IIf(cmbExecutor Is Nothing, "", cmbExecutor.Value), _
                          IIf(cmbDocumentType Is Nothing, "", _
                              IIf(cmbDocumentType.ListIndex >= 0, cmbDocumentType.Value, ""))
End Sub

'=====================================================================
'      AUTO-UPDATE ADDRESSES
'=====================================================================
Private Sub AutoUpdateAddressIfChanged()
    On Error Resume Next
    
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
'      MULTILINE FIELDS CONFIGURATION
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
'      SAFE FOCUS SETTING
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
'  CANCEL AND CLOSE BUTTONS
'=====================================================================
Private Sub btnCancel_Click()
    If MsgBox("Cancel letter creation?", vbYesNo + vbQuestion) = vbYes Then
        ClearCache
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If documentsList.count > 0 Then
        If MsgBox("Unsaved documents will be lost. Close?", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
        Else
            ClearCache
        End If
    Else
        ClearCache
    End If
End Sub

