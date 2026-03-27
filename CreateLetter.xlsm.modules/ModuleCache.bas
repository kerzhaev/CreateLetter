Attribute VB_Name = "ModuleCache"
' ======================================================================
' Module: ModuleCache
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Purpose: Caching system to accelerate searches
' Version: 1.4.2 - 27.03.2026
' ======================================================================

Option Explicit

Private addressCache As Object      ' Dictionary for addresses cache
Private attachmentCache As Object   ' Dictionary for attachments cache

Public Function GetCachedAddresses(searchTerm As String) As Collection
    On Error GoTo CacheError

    If addressCache Is Nothing Then
        Set addressCache = CreateObject("Scripting.Dictionary")
    End If

    Dim searchKey As String
    searchKey = UCase$(Trim$(searchTerm))

    If addressCache.Exists(searchKey) Then
        Set GetCachedAddresses = addressCache(searchKey)
        Debug.Print "Addresses loaded from cache: " & searchTerm
    Else
        Set GetCachedAddresses = SearchAddresses(searchTerm)

        If Not GetCachedAddresses Is Nothing Then
            addressCache.Add searchKey, GetCachedAddresses
            Debug.Print "Addresses loaded from database and cached: " & searchTerm & " (found: " & GetCachedAddresses.count & ")"
        Else
            Set GetCachedAddresses = New Collection
            Debug.Print "Addresses not found: " & searchTerm
        End If
    End If

    Exit Function

CacheError:
    Debug.Print "Addresses caching error: " & Err.description
    Set GetCachedAddresses = New Collection
End Function

Public Function GetCachedAttachments(searchTerm As String) As Collection
    On Error GoTo CacheError

    If attachmentCache Is Nothing Then
        Set attachmentCache = CreateObject("Scripting.Dictionary")
    End If

    Dim searchKey As String
    searchKey = UCase$(Trim$(searchTerm))

    If attachmentCache.Exists(searchKey) Then
        Set GetCachedAttachments = attachmentCache(searchKey)
        Debug.Print "Attachments loaded from cache: " & searchTerm
    Else
        Set GetCachedAttachments = SearchAttachments(searchTerm)

        If Not GetCachedAttachments Is Nothing Then
            attachmentCache.Add searchKey, GetCachedAttachments
            Debug.Print "Attachments loaded from database and cached: " & searchTerm & " (found: " & GetCachedAttachments.count & ")"
        Else
            Set GetCachedAttachments = New Collection
            Debug.Print "Attachments not found: " & searchTerm
        End If
    End If

    Exit Function

CacheError:
    Debug.Print "Attachments caching error: " & Err.description
    Set GetCachedAttachments = New Collection
End Function

Public Sub ClearCache()
    On Error Resume Next

    If Not addressCache Is Nothing Then
        addressCache.RemoveAll
        Set addressCache = Nothing
    End If

    If Not attachmentCache Is Nothing Then
        attachmentCache.RemoveAll
        Set attachmentCache = Nothing
    End If

    Debug.Print "Cache fully cleared"
    On Error GoTo 0
End Sub

Public Sub ClearAddressCache()
    On Error Resume Next

    If Not addressCache Is Nothing Then
        addressCache.RemoveAll
        Set addressCache = Nothing
    End If

    Debug.Print "Addresses cache cleared"
    On Error GoTo 0
End Sub

Public Sub ClearAttachmentCache()
    On Error Resume Next

    If Not attachmentCache Is Nothing Then
        attachmentCache.RemoveAll
        Set attachmentCache = Nothing
    End If

    Debug.Print "Attachments cache cleared"
    On Error GoTo 0
End Sub

Public Function GetCacheStats() As String
    Dim stats As String
    stats = "=== CACHE STATISTICS ===" & vbCrLf

    On Error Resume Next

    If addressCache Is Nothing Then
        stats = stats & "Addresses cache: not initialized" & vbCrLf
    Else
        stats = stats & "Addresses cache: " & addressCache.count & " entries" & vbCrLf
    End If

    If attachmentCache Is Nothing Then
        stats = stats & "Attachments cache: not initialized" & vbCrLf
    Else
        stats = stats & "Attachments cache: " & attachmentCache.count & " entries" & vbCrLf
    End If

    On Error GoTo 0

    GetCacheStats = stats
End Function

Public Sub ShowCacheStats()
    MsgBox GetCacheStats(), vbInformation, "Cache Statistics"
End Sub

