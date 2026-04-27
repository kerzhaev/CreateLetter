VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLetterHistoryRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class: clsLetterHistoryRecord
' Author: CreateLetter contributors
' Purpose: Typed DTO for letter history rows loaded from the Letters worksheet
' Version: 1.0.0 - 29.03.2026
' ======================================================================

Private pAddressee As String
Private pOutgoingNumber As String
Private pOutgoingDate As String
Private pAttachmentText As String
Private pDocumentSum As String
Private pReturnStatus As String
Private pExecutor As String
Private pDocumentTypeKey As String
Private pRowNumber As Long
Private pDispatchPackedFlag As String
Private pDispatchBatchId As String
Private pDispatchRegistryNumber As String
Private pDispatchRegistryDate As String

Public Property Get Addressee() As String
    Addressee = pAddressee
End Property

Public Property Let Addressee(ByVal value As String)
    pAddressee = value
End Property

Public Property Get OutgoingNumber() As String
    OutgoingNumber = pOutgoingNumber
End Property

Public Property Let OutgoingNumber(ByVal value As String)
    pOutgoingNumber = value
End Property

Public Property Get OutgoingDate() As String
    OutgoingDate = pOutgoingDate
End Property

Public Property Let OutgoingDate(ByVal value As String)
    pOutgoingDate = value
End Property

Public Property Get attachmentText() As String
    attachmentText = pAttachmentText
End Property

Public Property Let attachmentText(ByVal value As String)
    pAttachmentText = value
End Property

Public Property Get DocumentSum() As String
    DocumentSum = pDocumentSum
End Property

Public Property Let DocumentSum(ByVal value As String)
    pDocumentSum = value
End Property

Public Property Get returnStatus() As String
    returnStatus = pReturnStatus
End Property

Public Property Let returnStatus(ByVal value As String)
    pReturnStatus = value
End Property

Public Property Get Executor() As String
    Executor = pExecutor
End Property

Public Property Let Executor(ByVal value As String)
    pExecutor = value
End Property

Public Property Get DocumentTypeKey() As String
    DocumentTypeKey = pDocumentTypeKey
End Property

Public Property Let DocumentTypeKey(ByVal value As String)
    pDocumentTypeKey = value
End Property

Public Property Get rowNumber() As Long
    rowNumber = pRowNumber
End Property

Public Property Let rowNumber(ByVal value As Long)
    pRowNumber = value
End Property

Public Property Get DispatchPackedFlag() As String
    DispatchPackedFlag = pDispatchPackedFlag
End Property

Public Property Let DispatchPackedFlag(ByVal value As String)
    pDispatchPackedFlag = value
End Property

Public Property Get DispatchBatchId() As String
    DispatchBatchId = pDispatchBatchId
End Property

Public Property Let DispatchBatchId(ByVal value As String)
    pDispatchBatchId = value
End Property

Public Property Get DispatchRegistryNumber() As String
    DispatchRegistryNumber = pDispatchRegistryNumber
End Property

Public Property Let DispatchRegistryNumber(ByVal value As String)
    pDispatchRegistryNumber = value
End Property

Public Property Get DispatchRegistryDate() As String
    DispatchRegistryDate = pDispatchRegistryDate
End Property

Public Property Let DispatchRegistryDate(ByVal value As String)
    pDispatchRegistryDate = value
End Property

Public Function ToPartsArray() As Variant
    Dim parts(HistoryPartRowNumber) As String

    parts(HistoryPartAddressee) = pAddressee
    parts(HistoryPartOutgoingNumber) = pOutgoingNumber
    parts(HistoryPartOutgoingDate) = pOutgoingDate
    parts(HistoryPartAttachmentText) = pAttachmentText
    parts(HistoryPartDocumentSum) = pDocumentSum
    parts(HistoryPartReturnStatus) = pReturnStatus
    parts(HistoryPartExecutor) = pExecutor
    parts(HistoryPartDocumentType) = pDocumentTypeKey
    parts(HistoryPartRowNumber) = CStr(pRowNumber)

    ToPartsArray = parts
End Function
