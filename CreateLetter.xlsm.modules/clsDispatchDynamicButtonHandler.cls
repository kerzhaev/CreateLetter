VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDispatchDynamicButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class: clsDispatchDynamicButtonHandler
' Author: CreateLetter contributors
' Purpose: WithEvents bridge for runtime-created dispatch-form controls
' Version: 1.1.0 - 27.04.2026
' ======================================================================

Public WithEvents Button As MSForms.CommandButton
Public WithEvents SearchTextBox As MSForms.TextBox
Public WithEvents RuntimeListBox As MSForms.ListBox

Private hostForm As Object

Public Sub BindButton(targetButton As MSForms.CommandButton, ownerForm As Object)
    Set Button = targetButton
    Set hostForm = ownerForm
End Sub

Public Sub BindTextBox(targetTextBox As MSForms.TextBox, ownerForm As Object)
    Set SearchTextBox = targetTextBox
    Set hostForm = ownerForm
End Sub

Public Sub BindListBox(targetListBox As MSForms.ListBox, ownerForm As Object)
    Set RuntimeListBox = targetListBox
    Set hostForm = ownerForm
End Sub

Private Sub Button_Click()
    If hostForm Is Nothing Then Exit Sub
    hostForm.HandleDynamicButtonClick Button.Name
End Sub

Private Sub SearchTextBox_Change()
    If hostForm Is Nothing Then Exit Sub
    hostForm.HandleDynamicTextChanged SearchTextBox.Name
End Sub

Private Sub RuntimeListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Cancel = True
    If hostForm Is Nothing Then Exit Sub
    hostForm.HandleDynamicListDoubleClick RuntimeListBox.Name
End Sub
