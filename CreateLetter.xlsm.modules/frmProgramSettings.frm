VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgramSettings`n   Caption         =   "UserForm1"

   ClientHeight    =   3015

   ClientLeft      =   120

   ClientTop       =   465

   ClientWidth     =   4560

   OleObjectBlob   =   "frmProgramSettings.frx":0000

   StartUpPosition =   1  'CenterOwner

End

Attribute VB_Name = "frmProgramSettings"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False



' ======================================================================

' Form: frmProgramSettings v1.0.0

' Author: CreateLetter contributors

' Date: 03.05.2026

' Purpose: Runtime-built workbook-backed program settings UI

' ======================================================================

Option Explicit



Private WithEvents chkRequireOutgoingNumber As MSForms.CheckBox

Attribute chkRequireOutgoingNumber.VB_VarHelpID = -1

Private WithEvents btnSettingsSave As MSForms.CommandButton

Attribute btnSettingsSave.VB_VarHelpID = -1

Private WithEvents btnSettingsCancel As MSForms.CommandButton

Attribute btnSettingsCancel.VB_VarHelpID = -1



Private Sub UserForm_Initialize()



    On Error GoTo InitError



    Me.Caption = t("form.program_settings.title", "Program settings")

    Me.Width = 360

    Me.Height = 190

    Me.ScrollBars = fmScrollBarsNone



    BuildSettingsControls

    LoadSettingsValues

    FitSettingsFormToVisibleScreen

    Exit Sub



InitError:

    MsgBox t("form.program_settings.error.open_failed", "Failed to open program settings: ") & Err.description, vbExclamation, t("form.program_settings.title", "Program settings")



End Sub



Private Sub BuildSettingsControls()



    Dim introLabel As MSForms.Label

    Set introLabel = Me.Controls.Add("Forms.Label.1", "lblSettingsIntro", True)

    introLabel.Left = 12

    introLabel.Top = 12

    introLabel.Width = 320

    introLabel.Height = 36

    introLabel.WordWrap = True

    introLabel.Caption = t("form.program_settings.intro", "Shared settings are stored in this workbook and apply to all users.")



    Set chkRequireOutgoingNumber = Me.Controls.Add("Forms.CheckBox.1", "chkRequireOutgoingNumber", True)

    chkRequireOutgoingNumber.Left = 12

    chkRequireOutgoingNumber.Top = 62

    chkRequireOutgoingNumber.Width = 320

    chkRequireOutgoingNumber.Height = 36

    chkRequireOutgoingNumber.WordWrap = True

    chkRequireOutgoingNumber.Caption = t("form.program_settings.require_outgoing_number", "Require completed outgoing letter number")

    chkRequireOutgoingNumber.ControlTipText = t("form.program_settings.tip.require_outgoing_number", "When enabled, the letter form blocks navigation without a number after the slash, for example 7/125.")



    Set btnSettingsSave = Me.Controls.Add("Forms.CommandButton.1", "btnSettingsSave", True)

    btnSettingsSave.Left = 144

    btnSettingsSave.Top = 116

    btnSettingsSave.Width = 88

    btnSettingsSave.Height = 26

    btnSettingsSave.Caption = t("form.program_settings.save", "Save")



    Set btnSettingsCancel = Me.Controls.Add("Forms.CommandButton.1", "btnSettingsCancel", True)

    btnSettingsCancel.Left = 244

    btnSettingsCancel.Top = 116

    btnSettingsCancel.Width = 88

    btnSettingsCancel.Height = 26

    btnSettingsCancel.Caption = t("form.program_settings.cancel", "Cancel")



End Sub



Private Sub LoadSettingsValues()



    chkRequireOutgoingNumber.value = IsOutgoingNumberRequired()



End Sub



Private Sub btnSettingsSave_Click()



    On Error GoTo SaveError



    SetOutgoingNumberRequired CBool(chkRequireOutgoingNumber.value)

    MsgBox t("form.program_settings.saved", "Program settings saved in this workbook."), vbInformation, t("form.program_settings.title", "Program settings")

    Unload Me

    Exit Sub



SaveError:

    MsgBox t("form.program_settings.error.save_failed", "Failed to save program settings: ") & Err.description, vbExclamation, t("form.program_settings.title", "Program settings")



End Sub



Private Sub btnSettingsCancel_Click()



    Unload Me



End Sub



Private Sub FitSettingsFormToVisibleScreen()



    On Error GoTo FitError



    Dim maxWidth As Single

    Dim maxHeight As Single

    maxWidth = Application.UsableWidth - 20

    maxHeight = Application.UsableHeight - 20



    If maxWidth > 0 And Me.Width > maxWidth Then

        Me.Width = maxWidth

        Me.ScrollBars = fmScrollBarsBoth

        Me.ScrollWidth = 360

    End If



    If maxHeight > 0 And Me.Height > maxHeight Then

        Me.Height = maxHeight

        Me.ScrollBars = fmScrollBarsBoth

        Me.ScrollHeight = 190

    End If



    Exit Sub



FitError:

    Debug.Print "FitSettingsFormToVisibleScreen error: " & Err.description



End Sub
