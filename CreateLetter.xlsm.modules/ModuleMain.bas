Attribute VB_Name = "ModuleMain"
' ======================================================================
' Module: ModuleMain (main module) - WITH DEBUGGING
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Core shared logic for validation, data processing, Word generation, and workbook persistence
' Version: 1.6.2 — 26.03.2026
' ======================================================================

Option Explicit

' ======================================================================
'                    NEW FUNCTIONS v1.6.2
' ======================================================================

Public Function ValidateRequiredFields(addressee As String, city As String, region As String, postalCode As String, executor As String) As String
    If Len(Trim(addressee)) = 0 Then
        ValidateRequiredFields = "Fill in the 'Recipient Name' field"
        Exit Function
    End If
    
    If Len(Trim(city)) = 0 Then
        ValidateRequiredFields = "Fill in the 'City' field - this is a required field"
        Exit Function
    End If
    
    If Len(Trim(region)) = 0 Then
        ValidateRequiredFields = "Fill in the 'Region' field - this is a required field"
        Exit Function
    End If
    
    If Len(Trim(postalCode)) = 0 Then
        ValidateRequiredFields = "Fill in the 'Postal Code' field - this is a required field"
        Exit Function
    End If
    
    ValidateRequiredFields = ""
End Function

Public Function ValidateCreatorPage(pageIndex As Integer, addressee As String, city As String, region As String, postalCode As String, phoneNumber As String, letterNumber As String, letterDateText As String, executor As String, documentsCount As Long, ByRef focusControlName As String) As String
    focusControlName = ""
    ValidateCreatorPage = ""
    
    Select Case pageIndex
        Case 0
            If Len(Trim(addressee)) = 0 Then
                focusControlName = "txtAddressee"
                ValidateCreatorPage = "Fill in the 'Addressee' field."
                Exit Function
            End If
            
            If Len(Trim(city)) = 0 Then
                focusControlName = "txtCity"
                ValidateCreatorPage = "Fill in the 'City' field. This field is required."
                Exit Function
            End If
            
            If Len(Trim(region)) = 0 Then
                focusControlName = "txtRegion"
                ValidateCreatorPage = "Fill in the 'Region' field. This field is required."
                Exit Function
            End If
            
            If Len(Trim(postalCode)) = 0 Then
                focusControlName = "txtPostalCode"
                ValidateCreatorPage = "Fill in the 'Postal code' field. This field is required."
                Exit Function
            End If
            
            If Len(Trim(phoneNumber)) > 0 And Not IsPhoneNumberValid(phoneNumber) Then
                focusControlName = "txtAddresseePhone"
                ValidateCreatorPage = "Enter a valid addressee phone number."
                Exit Function
            End If
            
        Case 1
            If Len(Trim(letterNumber)) = 0 Then
                focusControlName = "txtLetterNumber"
                ValidateCreatorPage = "Enter the outgoing letter number."
                Exit Function
            End If
            
            If Len(Trim(letterDateText)) = 0 Then
                focusControlName = "txtLetterDate"
                ValidateCreatorPage = "Enter the letter date."
                Exit Function
            End If
            
            If Len(Trim(executor)) = 0 Then
                focusControlName = "cmbExecutor"
                ValidateCreatorPage = "Select an executor. This field is required."
                Exit Function
            End If
            
            Dim parsedDate As Date
            If Not TryParseDate(letterDateText, parsedDate) Then
                focusControlName = "txtLetterDate"
                ValidateCreatorPage = "Invalid letter date format."
                Exit Function
            End If
            
        Case 2
            If documentsCount = 0 Then
                ValidateCreatorPage = "Add at least one attachment document."
                Exit Function
            End If
    End Select
End Function

Public Function ValidateCreatorSubmission(addressee As String, city As String, region As String, postalCode As String, letterNumber As String, letterDateText As String, executor As String, documentsCount As Long, ByRef focusControlName As String) As String
    focusControlName = ""
    ValidateCreatorSubmission = ""
    
    If Len(Trim(addressee)) = 0 Then
        focusControlName = "txtAddressee"
        ValidateCreatorSubmission = "Addressee is not filled in."
        Exit Function
    End If
    
    If Len(Trim(city)) = 0 Then
        focusControlName = "txtCity"
        ValidateCreatorSubmission = "City is not filled in."
        Exit Function
    End If
    
    If Len(Trim(region)) = 0 Then
        focusControlName = "txtRegion"
        ValidateCreatorSubmission = "Region is not filled in."
        Exit Function
    End If
    
    If Len(Trim(postalCode)) = 0 Then
        focusControlName = "txtPostalCode"
        ValidateCreatorSubmission = "Postal code is not filled in."
        Exit Function
    End If
    
    If Len(Trim(letterNumber)) = 0 Then
        focusControlName = "txtLetterNumber"
        ValidateCreatorSubmission = "Letter number is not filled in."
        Exit Function
    End If
    
    If Len(Trim(letterDateText)) = 0 Then
        focusControlName = "txtLetterDate"
        ValidateCreatorSubmission = "Letter date is not filled in."
        Exit Function
    End If
    
    If Len(Trim(executor)) = 0 Then
        focusControlName = "cmbExecutor"
        ValidateCreatorSubmission = "Executor is not selected."
        Exit Function
    End If
    
    If documentsCount = 0 Then
        focusControlName = "txtAttachmentSearch"
        ValidateCreatorSubmission = "Add at least one document."
        Exit Function
    End If
End Function

Public Function FormatPhoneNumber(phoneInput As String) As String
    If Len(Trim(phoneInput)) = 0 Then
        FormatPhoneNumber = ""
        Exit Function
    End If
    
    Dim cleanPhone As String, i As Integer
    For i = 1 To Len(phoneInput)
        If IsNumeric(Mid(phoneInput, i, 1)) Then
            cleanPhone = cleanPhone & Mid(phoneInput, i, 1)
        End If
    Next i
    
    Select Case Len(cleanPhone)
        Case 11
            If Left(cleanPhone, 1) = "8" Or Left(cleanPhone, 1) = "7" Then
                FormatPhoneNumber = Left(cleanPhone, 1) & "-" & _
                                  Mid(cleanPhone, 2, 3) & "-" & _
                                  Mid(cleanPhone, 5, 3) & "-" & _
                                  Mid(cleanPhone, 8, 2) & "-" & _
                                  Mid(cleanPhone, 10, 2)
            Else
                FormatPhoneNumber = cleanPhone
            End If
            
        Case 10
            FormatPhoneNumber = "8-" & Left(cleanPhone, 3) & "-" & _
                              Mid(cleanPhone, 4, 3) & "-" & _
                              Mid(cleanPhone, 7, 2) & "-" & _
                              Mid(cleanPhone, 9, 2)
                              
        Case 7
            FormatPhoneNumber = Left(cleanPhone, 3) & "-" & _
                              Mid(cleanPhone, 4, 2) & "-" & _
                              Mid(cleanPhone, 6, 2)
                              
        Case Else
            FormatPhoneNumber = phoneInput
    End Select
End Function

Public Function IsPhoneNumberValid(phoneNumber As String) As Boolean
    Dim cleanPhone As String, i As Integer
    
    For i = 1 To Len(phoneNumber)
        If IsNumeric(Mid(phoneNumber, i, 1)) Then
            cleanPhone = cleanPhone & Mid(phoneNumber, i, 1)
        End If
    Next i
    
    IsPhoneNumberValid = (Len(cleanPhone) >= 7 And Len(cleanPhone) <= 11)
End Function

' ======================================================================
'                    DOCUMENT FUNCTIONS
' ======================================================================
Public Function CreateDocumentArray(docName As String, docNumber As String, docDate As String, docCopies As String, docSheets As String) As Variant
    Dim docArray(4) As String
    docArray(0) = Trim(docName)
    docArray(1) = Trim(docNumber)
    docArray(2) = Trim(docDate)
    docArray(3) = Trim(docCopies)
    docArray(4) = Trim(docSheets)
    
    CreateDocumentArray = docArray
End Function

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

Public Function FormatDocumentName(docArray As Variant) As String
    If Not IsArray(docArray) Then
        FormatDocumentName = "Error: invalid data format"
        Exit Function
    End If
    
    Dim result As String
    result = docArray(0)
    
    result = result & " No."
    If Len(Trim(docArray(1))) > 0 Then
        result = result & docArray(1)
    Else
        result = result & "    "
    End If
    
    result = result & " dated "
    If Len(Trim(docArray(2))) > 0 Then
        result = result & docArray(2)
    Else
        result = result & "        "
    End If
    
    result = result & " ("
    
    If Len(Trim(docArray(3))) > 0 Then
        result = result & docArray(3) & " copies"
    Else
        result = result & "  copies"
    End If
    
    result = result & ", "
    If Len(Trim(docArray(4))) > 0 Then
        result = result & docArray(4) & " sheets"
    Else
        result = result & "   sheets"
    End If
    
    result = result & ")"
    
    FormatDocumentName = result
End Function

' ======================================================================
'                    SEARCH AND DATA FUNCTIONS
' ======================================================================
Public Function SearchAddresses(searchTerm As String) As Collection
    Set SearchAddresses = New Collection
    
    On Error GoTo SearchError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim searchLine As String
        searchLine = ws.Cells(i, 1).Value & " " & _
                     ws.Cells(i, 2).Value & " " & _
                     ws.Cells(i, 3).Value & " " & _
                     ws.Cells(i, 4).Value & " " & _
                     ws.Cells(i, 5).Value & " " & _
                     ws.Cells(i, 6).Value & " " & _
                     ws.Cells(i, 7).Value
        
        If InStr(1, UCase(searchLine), UCase(searchTerm)) > 0 Then
            Dim fullAddress As String
            fullAddress = ws.Cells(i, 1).Value & " | " & _
                          ws.Cells(i, 2).Value & " | " & _
                          ws.Cells(i, 3).Value & " | " & _
                          ws.Cells(i, 4).Value & " | " & _
                          ws.Cells(i, 5).Value & " | " & _
                          ws.Cells(i, 6).Value & " | " & _
                          ws.Cells(i, 7).Value & " | " & i
            SearchAddresses.Add fullAddress
        End If
    Next i
    
    Exit Function
    
SearchError:
    Debug.Print "Address search error: " & Err.description
End Function

Public Function SearchAttachments(searchTerm As String) As Collection
    Set SearchAttachments = New Collection
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If Len(Trim(ws.Cells(i, 1).Value)) > 0 And InStr(1, UCase(ws.Cells(i, 1).Value), UCase(searchTerm)) > 0 Then
            SearchAttachments.Add ws.Cells(i, 1).Value
        End If
    Next i
End Function

' ======================================================================
'                    EXECUTOR FUNCTIONS
' ======================================================================
Public Function GetExecutorsList() As Collection
    Set GetExecutorsList = New Collection
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If Len(Trim(ws.Cells(i, 3).Value)) > 0 Then
            GetExecutorsList.Add ws.Cells(i, 3).Value
        End If
    Next i
End Function

Public Function GetCurrentUserFIO() As String
    On Error Resume Next
    GetCurrentUserFIO = Environ("USERNAME")
    If GetCurrentUserFIO = "" Then GetCurrentUserFIO = "Unknown user"
    On Error GoTo 0
End Function

Public Function GetExecutorPhone(executorFIO As String) As String
    GetExecutorPhone = "Not specified"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = executorFIO Then
            If Len(Trim(ws.Cells(i, 4).Value)) > 0 Then
                GetExecutorPhone = ws.Cells(i, 4).Value
            End If
            Exit Function
        End If
    Next i
End Function

' ======================================================================
'                    DATA SAVING FUNCTIONS
' ======================================================================
Public Sub SaveNewAddress(addressArray As Variant)
    On Error GoTo SaveError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    Dim i As Long
    For i = 0 To UBound(addressArray)
        If i = 6 Then
            ws.Cells(newRow, i + 1).Value = FormatPhoneNumber(CStr(addressArray(i)))
        Else
            ws.Cells(newRow, i + 1).Value = addressArray(i)
        End If
    Next i
    
    Exit Sub
    
SaveError:
    MsgBox "Error saving address: " & Err.description, vbCritical
End Sub

Public Sub UpdateExistingAddress(rowNumber As Long, addressArray As Variant)
    On Error GoTo UpdateError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    Dim i As Long
    For i = 0 To UBound(addressArray)
        If i = 6 Then
            ws.Cells(rowNumber, i + 1).Value = FormatPhoneNumber(CStr(addressArray(i)))
        Else
            ws.Cells(rowNumber, i + 1).Value = addressArray(i)
        End If
    Next i
    
    Exit Sub
    
UpdateError:
    MsgBox "Error updating address: " & Err.description, vbCritical
End Sub

Public Sub DeleteExistingAddress(rowNumber As Long)
    On Error GoTo DeleteError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    ws.Rows(rowNumber).Delete
    
    Exit Sub
    
DeleteError:
    MsgBox "Error deleting address: " & Err.description, vbCritical
End Sub

Public Function IsAddressDuplicate(addressArray As Variant, Optional excludeRow As Long = 0) As Boolean
    IsAddressDuplicate = False
    
    On Error GoTo CheckError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long, matchCount As Integer
    For i = 2 To lastRow
        If i = excludeRow Then GoTo NextRow
        
        matchCount = 0
        
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(addressArray(0))) Then matchCount = matchCount + 1
        If UCase(Trim(ws.Cells(i, 3).Value)) = UCase(Trim(addressArray(2))) Then matchCount = matchCount + 1
        If UCase(Trim(ws.Cells(i, 6).Value)) = UCase(Trim(addressArray(5))) Then matchCount = matchCount + 1
        
        If matchCount >= 3 Then
            IsAddressDuplicate = True
            Exit Function
        End If
        
NextRow:
    Next i
    
    Exit Function
    
CheckError:
    IsAddressDuplicate = False
End Function

' ======================================================================
'                    DEBUGGING FUNCTIONS
' ======================================================================

Public Sub SaveLetterInfoWithSum(addressee As String, letterNumber As String, letterDate As Date, documents As Collection, executor As String, documentType As String)
    ' === DEBUG START ===
    Debug.Print "=== DEBUG SaveLetterInfoWithSum START ==="
    Debug.Print "Addressee: " & addressee
    Debug.Print "LetterNumber: " & letterNumber
    Debug.Print "LetterDate: " & letterDate
    Debug.Print "Executor: " & executor
    Debug.Print "DocumentType: " & documentType
    Debug.Print "Documents count: " & documents.count
    
    Dim i As Long
    For i = 1 To documents.count
        Dim docArray As Variant
        docArray = documents(i)
        
        Debug.Print "Document #" & i & ": UBound=" & UBound(docArray) & " LBound=" & LBound(docArray)
        
        Dim j As Long
        For j = LBound(docArray) To UBound(docArray)
            Debug.Print "  Element " & j & ": '" & CStr(docArray(j)) & "'"
        Next j
    Next i
    Debug.Print "=== DEBUG SaveLetterInfoWithSum INITIAL END ==="
    ' === DEBUG END ===
    
    On Error GoTo SaveLetterError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Letters")
    
    Dim newRow As Long
    newRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    ' DEBUG: Writing basic data
    Debug.Print "=== BEFORE writing basic data ==="
    ws.Cells(newRow, 1).Value = addressee
    ws.Cells(newRow, 2).Value = letterNumber
    ws.Cells(newRow, 3).Value = letterDate
    Debug.Print "=== AFTER writing basic data ==="
    
    ' DEBUG: Formatting attachments
    Debug.Print "=== BEFORE FormatAttachmentsListCompactWithSum ==="
    Dim attachmentText As String
    attachmentText = FormatAttachmentsListCompactWithSum(documents)
    Debug.Print "=== AFTER FormatAttachmentsListCompactWithSum ==="
    Debug.Print "Result length: " & Len(attachmentText)
    Debug.Print "Result preview: " & Left(attachmentText, 200)
    
    ' DEBUG: Writing to Excel
    Debug.Print "=== BEFORE writing to Excel cell (4) ==="
    On Error Resume Next
    ws.Cells(newRow, 4).Value = attachmentText
    If Err.number <> 0 Then
        Debug.Print "ERROR writing to cell (4): " & Err.description & " (Number: " & Err.number & ")"
        Err.Clear
    End If
    On Error GoTo SaveLetterError
    Debug.Print "=== AFTER writing to Excel cell (4) ==="
    
    ' DEBUG: Sum calculation
    Debug.Print "=== BEFORE CalculateTotalDocumentsSum ==="
    Dim totalSum As Double
    totalSum = CalculateTotalDocumentsSum(documents)
    Debug.Print "=== AFTER CalculateTotalDocumentsSum, result: " & totalSum
    
    ' DEBUG: Writing sum
    Debug.Print "=== BEFORE writing sum to cell (5) ==="
    On Error Resume Next
    If totalSum > 0 Then
        ws.Cells(newRow, 5).Value = totalSum
        Debug.Print "Written totalSum: " & totalSum
    Else
        ws.Cells(newRow, 5).Value = ""
        Debug.Print "Written empty sum"
    End If
    If Err.number <> 0 Then
        Debug.Print "ERROR writing to cell (5): " & Err.description & " (Number: " & Err.number & ")"
        Err.Clear
    End If
    On Error GoTo SaveLetterError
    Debug.Print "=== AFTER writing sum to cell (5) ==="
    
    ' DEBUG: Writing remaining data
    Debug.Print "=== BEFORE writing remaining cells ==="
    On Error Resume Next
    ws.Cells(newRow, 6).Value = ""
    ws.Cells(newRow, 7).Value = executor
    ws.Cells(newRow, 8).Value = documentType
    If Err.number <> 0 Then
        Debug.Print "ERROR writing remaining cells: " & Err.description & " (Number: " & Err.number & ")"
        Err.Clear
    End If
    On Error GoTo SaveLetterError
    Debug.Print "=== AFTER writing remaining cells ==="
    
    Debug.Print "=== DEBUG SaveLetterInfoWithSum SUCCESS END ==="
    
    Exit Sub
    
SaveLetterError:
    Debug.Print "=== ERROR in SaveLetterInfoWithSum ==="
    Debug.Print "Error Number: " & Err.number
    Debug.Print "Error Description: " & Err.description
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "==========================="
    MsgBox "Error saving letter info: " & Err.description, vbCritical
End Sub

Public Function FormatAttachmentsListCompactWithSum(documentsList As Collection) As String
    Debug.Print "=== DEBUG FormatAttachmentsListCompactWithSum START ==="
    
    If documentsList Is Nothing Or documentsList.count = 0 Then
        FormatAttachmentsListCompactWithSum = "Documents not specified"
        Debug.Print "=== DEBUG FormatAttachmentsListCompactWithSum END (empty) ==="
        Exit Function
    End If
    
    Dim result As String
    Dim i As Long
    
    For i = 1 To documentsList.count
        Debug.Print "Processing document " & i & " of " & documentsList.count
        
        If i > 1 Then result = result & "; "
        
        Dim docArray As Variant
        docArray = documentsList(i)
        
        If IsArray(docArray) And UBound(docArray) >= 5 Then
            Debug.Print "  Calling FormatDocumentNameWithSum"
            Dim docResult As String
            docResult = FormatDocumentNameWithSum(docArray)
            result = result & docResult & ";"
            Debug.Print "  Result so far: " & result
        Else
            Debug.Print "  Calling FormatDocumentName"
            result = result & FormatDocumentName(docArray) & ";"
            Debug.Print "  Result so far: " & result
        End If
    Next i
    
    Debug.Print "Final result length: " & Len(result)
    Debug.Print "Final result: " & result
    Debug.Print "=== DEBUG FormatAttachmentsListCompactWithSum END ==="
    
    FormatAttachmentsListCompactWithSum = result
End Function

Public Function FormatDocumentNameWithSum(docArray As Variant) As String
    Debug.Print "=== DEBUG FormatDocumentNameWithSum START ==="
    Debug.Print "IsArray: " & IsArray(docArray)
    
    If Not IsArray(docArray) Then
        FormatDocumentNameWithSum = "Error: invalid data format"
        Debug.Print "ERROR: Not array"
        Debug.Print "=== DEBUG FormatDocumentNameWithSum END ==="
        Exit Function
    End If
    
    Debug.Print "Array UBound: " & UBound(docArray) & " LBound: " & LBound(docArray)
    
    Dim j As Long
    For j = LBound(docArray) To UBound(docArray)
        Debug.Print "  Element " & j & ": '" & CStr(docArray(j)) & "'"
    Next j
    
    Dim result As String
    result = docArray(0)
    
    result = result & " No."
    If Len(Trim(docArray(1))) > 0 Then
        result = result & docArray(1)
    Else
        result = result & "    "
    End If
    
    result = result & " dated "
    If Len(Trim(docArray(2))) > 0 Then
        result = result & docArray(2)
    Else
        result = result & "        "
    End If
    
    ' FIXED SUM CHECK
    If UBound(docArray) >= 5 And Len(Trim(docArray(5))) > 0 Then
        Debug.Print "Processing sum: '" & docArray(5) & "'"
        If IsNumeric(docArray(5)) Then
            Dim sumText As String
            sumText = CStr(CLng(CDbl(docArray(5))))
            result = result & " for the amount of " & sumText & " rub."
            Debug.Print "Sum formatted as: " & sumText
        Else
            result = result & " (" & docArray(5) & ")"
            Debug.Print "Sum as text: " & docArray(5)
        End If
    Else
        Debug.Print "No sum found or empty sum"
    End If
    
    result = result & " ("
    
    If Len(Trim(docArray(3))) > 0 Then
        result = result & docArray(3) & " copies"
    Else
        result = result & "  copies"
    End If
    
    result = result & ", "
    If Len(Trim(docArray(4))) > 0 Then
        result = result & docArray(4) & " sheets"
    Else
        result = result & "   sheets"
    End If
    
    result = result & ")"
    
    Debug.Print "Final document result: " & result
    Debug.Print "=== DEBUG FormatDocumentNameWithSum END ==="
    
    FormatDocumentNameWithSum = result
End Function

Public Function FormatAttachmentsListForWordWithSum(documentsList As Collection) As Collection
    Set FormatAttachmentsListForWordWithSum = New Collection
    
    If documentsList Is Nothing Or documentsList.count = 0 Then
        FormatAttachmentsListForWordWithSum.Add "documents not specified;"
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

Public Function BuildSummaryAttachmentsText(documentsList As Collection) As String
    If documentsList Is Nothing Or documentsList.count = 0 Then
        BuildSummaryAttachmentsText = ""
        Exit Function
    End If
    
    Dim attachmentText As String
    Dim i As Long
    
    For i = 1 To documentsList.count
        If i > 1 Then attachmentText = attachmentText & vbCrLf
        attachmentText = attachmentText & i & ". " & FormatDocumentNameWithSum(documentsList(i)) & ";"
    Next i
    
    BuildSummaryAttachmentsText = attachmentText
End Function

Public Function GetDocumentDisplayItems(documentsList As Collection) As Collection
    Set GetDocumentDisplayItems = New Collection
    
    If documentsList Is Nothing Then Exit Function
    
    Dim i As Long
    For i = 1 To documentsList.count
        GetDocumentDisplayItems.Add FormatDocumentNameWithSum(documentsList(i))
    Next i
End Function

Public Function DuplicateDocumentArray(sourceItem As Variant) As Variant
    Dim sourceName As String
    Dim sourceDate As String
    Dim sourceCopies As String
    Dim sourceSheets As String
    Dim sourceSum As String
    
    sourceName = ""
    sourceDate = ""
    sourceCopies = ""
    sourceSheets = ""
    sourceSum = ""
    
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
    
    DuplicateDocumentArray = CreateDocumentArrayWithSum(sourceName, "", sourceDate, sourceCopies, sourceSheets, sourceSum)
End Function

Public Sub MoveDocumentCollectionItemUp(documentsList As Collection, oneBasedIndex As Long)
    If documentsList Is Nothing Then Exit Sub
    If oneBasedIndex <= 1 Or oneBasedIndex > documentsList.count Then Exit Sub
    
    Dim tempDoc As Variant
    tempDoc = documentsList(oneBasedIndex - 1)
    documentsList.Remove oneBasedIndex - 1
    documentsList.Add tempDoc, , oneBasedIndex - 1
End Sub

Public Sub MoveDocumentCollectionItemDown(documentsList As Collection, oneBasedIndex As Long)
    If documentsList Is Nothing Then Exit Sub
    If oneBasedIndex < 1 Or oneBasedIndex >= documentsList.count Then Exit Sub
    
    Dim tempDoc As Variant
    tempDoc = documentsList(oneBasedIndex + 1)
    documentsList.Remove oneBasedIndex + 1
    documentsList.Add tempDoc, , oneBasedIndex
End Sub

Public Sub CreateLetterDocument(addressee As String, addressArray As Variant, letterNumber As String, letterDateRaw As String, executor As String, documentType As String, useAlternateTemplate As Boolean, documentsList As Collection)
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
        Err.Raise 429, "CreateLetterDocument", "Failed to create Word.Application object"
    End If
    
    wordApp.Visible = True
    
    Dim templatePath As String
    templatePath = GetLetterTemplatePath(useAlternateTemplate)
    
    If Dir(templatePath) <> "" Then
        Set wordDoc = wordApp.Documents.Open(templatePath)
        If Not wordDoc Is Nothing Then
            FillWordTemplateData wordDoc, addressee, addressArray, letterNumber, letterDateRaw, executor, documentType, documentsList
            GoTo SaveDocument
        End If
    End If
    
    Set wordDoc = wordApp.Documents.Add
    CreateLetterDocumentFromScratch wordDoc, addressee, addressArray, letterNumber, letterDateRaw, executor, documentType, documentsList
    
SaveDocument:
    Dim fileName As String
    fileName = GenerateFileNameWithExecutor(IIf(Len(Trim(addressee)) = 0, "Letter", addressee), letterNumber, executor)
    
    wordDoc.SaveAs fileName
    Debug.Print "File saved: " & fileName
    
    On Error Resume Next
    ThisWorkbook.Save
    Debug.Print "Excel workbook saved"
    On Error GoTo ErrorHandler
    
    wordApp.Visible = True
    wordDoc.Activate
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating letter: " & Err.Description, vbCritical
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Public Sub FillWordTemplateData(wordDoc As Object, addresseeText As String, addressArray As Variant, numberText As String, rawDateText As String, executorText As String, documentType As String, documentsList As Collection)
    On Error GoTo TemplateError
    
    Dim addressText As String
    Dim dateText As String
    Dim phoneText As String
    Dim letterText As String
    
    addressText = FormatRecipientAddress(addressArray)
    dateText = FormatLetterDate(rawDateText)
    Debug.Print "Formatted date: " & dateText
    
    phoneText = GetExecutorPhone(executorText)
    letterText = GetDocumentTypeText(documentType)
    
    SafeReplaceInWord wordDoc, "RecipientName", addresseeText
    SafeReplaceInWord wordDoc, "RecipientAddress", addressText
    SafeReplaceInWord wordDoc, "OutgoingNumber", numberText
    SafeReplaceInWord wordDoc, "OutgoingDate", dateText
    SafeReplaceInWord wordDoc, "ExecutorName", executorText
    SafeReplaceInWord wordDoc, "ExecutorPhone", phoneText
    SafeReplaceInWord wordDoc, "LetterText", letterText
    
    ReplaceAttachmentsInTemplateWithFontAndSum wordDoc, documentsList, 10
    Exit Sub
    
TemplateError:
    MsgBox "Template filling error: " & Err.Description, vbCritical
End Sub

Public Sub CreateLetterDocumentFromScratch(wordDoc As Object, addresseeText As String, addressArray As Variant, numberText As String, rawDateText As String, executorText As String, documentType As String, documentsList As Collection)
    On Error GoTo ScratchError
    
    Dim content As String
    Dim addressText As String
    Dim letterText As String
    Dim dateText As String
    
    addressText = FormatRecipientAddress(addressArray)
    letterText = GetDocumentTypeText(documentType)
    dateText = FormatLetterDate(rawDateText)
    
    content = "To the Commander of military unit " & addresseeText & vbCrLf & vbCrLf
    content = content & addressText & vbCrLf & vbCrLf & vbCrLf
    content = content & letterText & vbCrLf & vbCrLf
    content = content & "Executor: " & executorText & vbCrLf
    content = content & "Phone: " & GetExecutorPhone(executorText) & vbCrLf
    content = content & "Ref. No.: " & numberText & vbCrLf
    content = content & "Date: " & dateText & vbCrLf & vbCrLf
    
    wordDoc.Content.Text = content
    AppendAttachmentsToDocumentWithFontAndSum wordDoc, documentsList, 10
    Exit Sub
    
ScratchError:
    MsgBox "Letter creation error: " & Err.Description, vbCritical
End Sub

Public Sub ReplaceAttachmentsInTemplateWithFontAndSum(wordDoc As Object, documentsList As Collection, fontSize As Integer)
    On Error Resume Next
    
    Dim rng As Object
    Set rng = wordDoc.content
    
    With rng.Find
        .ClearFormatting
        .Forward = True
        .Wrap = 1
        .Text = "AttachmentsList"
        
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

Public Sub AppendAttachmentsToDocumentWithFontAndSum(wordDoc As Object, documentsList As Collection, fontSize As Integer)
    On Error Resume Next
    
    Dim rng As Object
    Set rng = wordDoc.content
    rng.Collapse 0
    
    rng.InsertAfter "Attachment: "
    
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

Public Function CalculateTotalDocumentsSum(documents As Collection) As Double
    Debug.Print "=== DEBUG CalculateTotalDocumentsSum START ==="
    CalculateTotalDocumentsSum = 0
    
    If documents Is Nothing Or documents.count = 0 Then
        Debug.Print "=== DEBUG CalculateTotalDocumentsSum END (empty) ==="
        Exit Function
    End If
    
    Dim documentsWithSum As Integer
    Dim totalSum As Double
    documentsWithSum = 0
    totalSum = 0
    
    Dim i As Long
    For i = 1 To documents.count
        Dim docArray As Variant
        docArray = documents(i)
        
        Debug.Print "Checking document " & i & " for sum calculation"
        
        If IsArray(docArray) And UBound(docArray) >= 5 Then
            Dim docSum As String
            docSum = Trim(CStr(docArray(5)))
            
            Debug.Print "  Document " & i & " sum: '" & docSum & "'"
            Debug.Print "  IsNumeric: " & IsNumeric(docSum)
            Debug.Print "  Len > 0: " & (Len(docSum) > 0)
            
            ' FIXED: Breaking down condition into nested IFs to avoid premature CDbl call
            If Len(docSum) > 0 Then
                If IsNumeric(docSum) Then
                    Dim sumValue As Double
                    sumValue = CDbl(docSum)
                    Debug.Print "  Converted sum: " & sumValue
                    
                    If sumValue > 0 Then
                        documentsWithSum = documentsWithSum + 1
                        totalSum = totalSum + sumValue
                        Debug.Print "  Added to total: " & sumValue
                    End If
                End If
            End If
        End If
    Next i
    
    If documentsWithSum > 1 Then
        CalculateTotalDocumentsSum = 0
        Debug.Print "Multiple documents with sum - returning 0"
    ElseIf documentsWithSum = 1 Then
        CalculateTotalDocumentsSum = totalSum
        Debug.Print "Single document with sum: " & totalSum
    Else
        CalculateTotalDocumentsSum = 0
        Debug.Print "No documents with sum - returning 0"
    End If
    
    Debug.Print "=== DEBUG CalculateTotalDocumentsSum END, result: " & CalculateTotalDocumentsSum
End Function


' ======================================================================
'                    REMAINING FUNCTIONS (abbreviated)
' ======================================================================

Public Function FormatRecipientAddress(addressParts As Variant) As String
    Dim fullAddress As String
    Dim addressComponents As Collection
    Set addressComponents = New Collection
    
    Dim i As Integer
    For i = 1 To UBound(addressParts)
        If Len(Trim(CStr(addressParts(i)))) > 0 Then
            addressComponents.Add Trim(CStr(addressParts(i)))
        End If
    Next i
    
    For i = 1 To addressComponents.count
        If i > 1 Then fullAddress = fullAddress & ", "
        fullAddress = fullAddress & addressComponents(i)
    Next i
    
    FormatRecipientAddress = fullAddress
End Function

Public Function TryParseDate(rawText As String, ByRef outDate As Date) As Boolean
    Dim t As String, ok As Boolean
    Dim clean As String, i As Long, ch As String
    
    On Error Resume Next
    If TryParseDateExtended(rawText, outDate) Then
        TryParseDate = True
        Exit Function
    End If
    On Error GoTo 0
    
    TryParseDate = False
    
    If Len(Trim(rawText)) = 0 Then Exit Function
    
    On Error Resume Next
    If IsDate(rawText) Then
        outDate = CDate(rawText)
        TryParseDate = True
        Exit Function
    End If
    On Error GoTo 0
    
    t = Replace(rawText, "/", ".")
    
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

Public Function ResolveLetterDateOrToday(rawText As String) As Date
    If Len(Trim(rawText)) = 0 Then
        ResolveLetterDateOrToday = Date
        Exit Function
    End If
    
    If IsDate(rawText) Then
        ResolveLetterDateOrToday = CDate(rawText)
        Exit Function
    End If
    
    Dim parsedDate As Date
    If TryParseDate(rawText, parsedDate) Then
        ResolveLetterDateOrToday = parsedDate
    Else
        ResolveLetterDateOrToday = Date
    End If
End Function

Public Function HasAddressDataChanged(rowNumber As Long, newAddressArray As Variant) As Boolean
    HasAddressDataChanged = False
    
    On Error GoTo CompareError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    Dim i As Long
    For i = 0 To UBound(newAddressArray)
        Dim sheetValue As String
        Dim formValue As String
        
        sheetValue = UCase(Trim(CStr(ws.Cells(rowNumber, i + 1).Value)))
        formValue = UCase(Trim(CStr(newAddressArray(i))))
        
        If sheetValue <> formValue Then
            Debug.Print "Change in column " & (i + 1) & ": '" & ws.Cells(rowNumber, i + 1).Value & "' -> '" & newAddressArray(i) & "'"
            HasAddressDataChanged = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
CompareError:
    Debug.Print "Error comparing address data: " & Err.Description
    HasAddressDataChanged = False
End Function

Public Function FormatLetterDate(dateValue As String) As String
    On Error GoTo FormatError
    
    Dim d As Date
    
    If IsDate(dateValue) Then
        d = CDate(dateValue)
    Else
        If TryParseDateExtended(dateValue, d) Then
        Else
            FormatLetterDate = dateValue
            Exit Function
        End If
    End If
    
    Dim dayNum As Integer, monthNum As Integer, yearNum As Integer
    dayNum = Day(d)
    monthNum = Month(d)
    yearNum = Year(d)
    
    Dim monthName As String
    monthName = GetDirectMonthName(monthNum)
    
    FormatLetterDate = dayNum & " " & monthName & " " & yearNum
    
    Exit Function
    
FormatError:
    FormatLetterDate = dateValue
End Function

Private Function GetDirectMonthName(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetDirectMonthName = BuildUnicodeString(&H44F, &H43D, &H432, &H430, &H440, &H44F)
        Case 2: GetDirectMonthName = BuildUnicodeString(&H444, &H435, &H432, &H440, &H430, &H43B, &H44F)
        Case 3: GetDirectMonthName = BuildUnicodeString(&H43C, &H430, &H440, &H442, &H430)
        Case 4: GetDirectMonthName = BuildUnicodeString(&H430, &H43F, &H440, &H435, &H43B, &H44F)
        Case 5: GetDirectMonthName = BuildUnicodeString(&H43C, &H430, &H44F)
        Case 6: GetDirectMonthName = BuildUnicodeString(&H438, &H44E, &H43D, &H44F)
        Case 7: GetDirectMonthName = BuildUnicodeString(&H438, &H44E, &H43B, &H44F)
        Case 8: GetDirectMonthName = BuildUnicodeString(&H430, &H432, &H433, &H443, &H441, &H442, &H430)
        Case 9: GetDirectMonthName = BuildUnicodeString(&H441, &H435, &H43D, &H442, &H44F, &H431, &H440, &H44F)
        Case 10: GetDirectMonthName = BuildUnicodeString(&H43E, &H43A, &H442, &H44F, &H431, &H440, &H44F)
        Case 11: GetDirectMonthName = BuildUnicodeString(&H43D, &H43E, &H44F, &H431, &H440, &H44F)
        Case 12: GetDirectMonthName = BuildUnicodeString(&H434, &H435, &H43A, &H430, &H431, &H440, &H44F)
        Case Else: GetDirectMonthName = "unknown_month"
    End Select
End Function

Private Function BuildUnicodeString(ParamArray codePoints() As Variant) As String
    Dim i As Long

    BuildUnicodeString = ""
    For i = LBound(codePoints) To UBound(codePoints)
        BuildUnicodeString = BuildUnicodeString & ChrW(CLng(codePoints(i)))
    Next i
End Function

Public Sub ShowLetterCreatorDelayed()
    On Error GoTo DelayedErrorHandler
    
    Load frmLetterCreator
    frmLetterCreator.Show vbModeless
    Exit Sub
    
DelayedErrorHandler:
    MsgBox "Failed to open letter creation form: " & Err.description, vbCritical
End Sub

Public Sub StartFormirovanieLetters()
    Load frmLetterCreator
    frmLetterCreator.Show vbModeless
End Sub

Public Function GetDocumentTypeText(documentType As String) As String
    GetDocumentTypeText = "forwarding confirmed accounting documents to your address"
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    
    If ws Is Nothing Then Exit Function
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Text")
    
    If Not tbl Is Nothing Then
        If tbl.ListRows.count >= 1 Then
            Dim textResult As String
            If UCase(Trim(documentType)) = "OWN FOR CONFIRMATION" Then
                textResult = Trim(tbl.DataBodyRange.Cells(1, 1).Value)
            ElseIf tbl.ListRows.count >= 2 Then
                textResult = Trim(tbl.DataBodyRange.Cells(2, 1).Value)
            End If
            
            If Len(textResult) > 0 Then
                textResult = LCase(Left(textResult, 1)) & Mid(textResult, 2)
                GetDocumentTypeText = textResult
            End If
        End If
    End If
    
    On Error GoTo 0
End Function

Public Sub SafeReplaceInWord(wordDoc As Object, findText As String, replaceText As String)
    On Error GoTo ReplaceError
    
    If findText = "ExecutorPhone" Then
        Debug.Print "Attempting to replace ExecutorPhone with: '" & replaceText & "'"
        Debug.Print "Replacement text length: " & Len(replaceText)
    End If
    
    If Len(replaceText) > 180 Then
        Dim fragments As Collection
        Set fragments = SplitStringToFragments(replaceText, 180)
        SafeReplaceInWordWithFragments wordDoc, findText, fragments
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
        
        If findText = "ExecutorPhone" Then
            Debug.Print "Replacement done. Checking result..."
            On Error Resume Next
            Dim rng As Object
            Set rng = wordDoc.content
            With rng.Find
                .Text = replaceText
                If .Execute Then
                    rng.Font.Name = "Times New Roman"
                    rng.Font.Size = 12
                    rng.Font.Color = RGB(0, 0, 0)
                    Debug.Print "Phone formatting applied"
                End If
            End With
            On Error GoTo ReplaceError
        End If
    End If
    
    Exit Sub
    
ReplaceError:
    Debug.Print "Error replacing '" & findText & "': " & Err.description
    On Error Resume Next
    wordDoc.content.Text = Replace(wordDoc.content.Text, findText, Left(replaceText, 200))
    On Error GoTo 0
End Sub

Public Sub SafeReplaceInWordWithFragments(wordDoc As Object, findText As String, fragments As Collection)
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
            
            Dim i As Long, fullText As String
            For i = 1 To fragments.count
                If i > 1 Then fullText = fullText & " "
                fullText = fullText & CStr(fragments(i))
            Next i
            
            rng.InsertAfter fullText
            
            On Error Resume Next
            Dim insertedRange As Object
            Set insertedRange = wordDoc.Range(rng.Start, rng.Start + Len(fullText))
            
            With insertedRange
                .Font.Name = "Times New Roman"
                .Font.Size = 12
                .ParagraphFormat.SpaceAfter = 0
                .ParagraphFormat.SpaceBefore = 0
                .ParagraphFormat.LineSpacing = 14
            End With
            On Error GoTo ReplaceError
        End If
    End With
    
    Exit Sub
    
ReplaceError:
    Debug.Print "Error replacing fragments '" & findText & "': " & Err.description
End Sub

Private Function SplitStringToFragments(inputString As String, maxLength As Integer) As Collection
    Set SplitStringToFragments = New Collection
    
    If Len(inputString) = 0 Then Exit Function
    
    Dim currentPos As Long, fragmentLength As Long
    Dim fragment As String
    currentPos = 1
    
    While currentPos <= Len(inputString)
        If currentPos + maxLength - 1 > Len(inputString) Then
            fragment = Mid(inputString, currentPos)
            SplitStringToFragments.Add fragment
            Exit Function
        End If
        
        fragment = Mid(inputString, currentPos, maxLength)
        
        Dim breakPos As Long
        breakPos = FindBestBreakPosition(fragment)
        
        If breakPos > 0 And breakPos < maxLength Then
            fragment = Mid(inputString, currentPos, breakPos)
            currentPos = currentPos + breakPos + 1
            
            While currentPos <= Len(inputString) And Mid(inputString, currentPos, 1) = " "
                currentPos = currentPos + 1
            Wend
        Else
            currentPos = currentPos + maxLength
        End If
        
        SplitStringToFragments.Add Trim(fragment)
    Wend
End Function

Private Function FindBestBreakPosition(textFragment As String) As Long
    Dim i As Long, testPos As Long
    
    For i = Len(textFragment) To Len(textFragment) \ 2 Step -1
        Dim currentChar As String
        currentChar = Mid(textFragment, i, 1)
        
        If currentChar = "." And i < Len(textFragment) Then
            If Mid(textFragment, i + 1, 1) = " " Then
                FindBestBreakPosition = i
                Exit Function
            End If
        End If
        
        If currentChar = "," And i < Len(textFragment) Then
            If Mid(textFragment, i + 1, 1) = " " Then
                testPos = i
            End If
        End If
        
        If currentChar = ":" And i < Len(textFragment) Then
            If Mid(textFragment, i + 1, 1) = " " Then
                If testPos = 0 Then testPos = i
            End If
        End If
        
        If currentChar = " " And testPos = 0 Then
            testPos = i - 1
        End If
    Next i
    
    FindBestBreakPosition = testPos
End Function

Public Sub FormatAttachmentsInWord(rng As Object, Optional fontSize As Integer = 10)
    On Error Resume Next
    
    rng.Font.Size = fontSize
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.LineSpacing = fontSize + 2
    
    On Error GoTo 0
End Sub

Public Function GenerateFileNameWithExecutor(addressee As String, letterNumber As String, executor As String) As String
    Dim cleanAddressee As String, cleanNumber As String, cleanExecutor As String
    Dim currentDate As String
    
    cleanAddressee = CleanFileName(addressee)
    cleanNumber = CleanFileName(letterNumber)
    cleanExecutor = CleanFileName(executor)
    currentDate = Format(Date, "dd.mm.yyyy")
    
    GenerateFileNameWithExecutor = ThisWorkbook.Path & "\" & cleanAddressee & "_" & _
                                  cleanNumber & "_" & currentDate & "_" & cleanExecutor & ".docx"
End Function

Private Function GetLetterTemplatePath(useAlternateTemplate As Boolean) As String
    If useAlternateTemplate Then
        GetLetterTemplatePath = ThisWorkbook.Path & "\LetterTemplateFOU.docx"
    Else
        GetLetterTemplatePath = ThisWorkbook.Path & "\LetterTemplate.docx"
    End If
End Function

Public Function CleanFileName(inputName As String) As String
    Dim result As String
    result = Trim(inputName)
    
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    result = Replace(result, " ", "_")
    
    If Len(result) > 30 Then result = Left(result, 30)
    
    CleanFileName = result
End Function

Public Sub ClearHighlight()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not ws Is Nothing Then
        ws.Cells.Interior.Pattern = xlNone
        ws.Cells.Interior.ColorIndex = xlNone
        Debug.Print "Row highlight cleared"
        
        Application.StatusBar = False
    End If
    
    On Error GoTo 0
End Sub

Public Sub RestoreFocusToHistory()
    On Error Resume Next
    
    Dim historyForm As Object
    Set historyForm = VBA.UserForms("frmLetterHistory")
    
    If Not historyForm Is Nothing Then
        historyForm.SetFocus
        historyForm.ZOrder 0
        Debug.Print "Focus returned to letter history form"
    End If
    
    On Error GoTo 0
End Sub

Public Sub ShowLetterHistoryModeless()
    On Error GoTo ShowHistoryError
    
    Dim existingForm As Object
    On Error Resume Next
    Set existingForm = VBA.UserForms("frmLetterHistory")
    On Error GoTo ShowHistoryError
    
    If Not existingForm Is Nothing Then
        existingForm.SetFocus
        existingForm.ZOrder 0
        MsgBox "Letter history form is already open!", vbInformation
    Else
        Load frmLetterHistory
        frmLetterHistory.Show vbModeless
        Debug.Print "Letter history form launched modelessly from ModuleMain"
    End If
    
    Exit Sub
    
ShowHistoryError:
    MsgBox "Error opening letter history form: " & Err.description, vbCritical
End Sub

Public Sub ClearAddressHighlight()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Addresses")
    
    If Not ws Is Nothing Then
        ws.Cells.Interior.Pattern = xlNone
        Debug.Print "Address row highlight cleared"
    End If
    
    On Error GoTo 0
End Sub

Public Sub ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = False
    Debug.Print "Excel status bar cleared"
    On Error GoTo 0
End Sub

Public Sub SetStatusBarMessage(message As String, Optional clearAfterSeconds As Integer = 0)
    On Error Resume Next
    
    Application.StatusBar = message
    Debug.Print "Status bar: " & message
    
    If clearAfterSeconds > 0 Then
        Application.OnTime Now + TimeValue("00:00:" & Format(clearAfterSeconds, "00")), "ClearStatusBar"
    End If
    
    On Error GoTo 0
End Sub

