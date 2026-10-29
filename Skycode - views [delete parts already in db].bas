Option Compare Database
Option Explicit

' ============================================================
' Skycode - Delete Extraneous Views
' Deletes views from the active Composer document where view names
' match a user-pasted list of values.
' ============================================================

Sub DeleteExtraneousViews()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim logFile As Object
    Dim fs As Object

    Dim rawInput As String
    Dim targetMap As Object
    Dim foundMap As Object

    Dim viewSelection As String
    Dim viewNames() As String
    Dim i As Long
    Dim key As String

    Dim requestedCount As Long
    Dim matchedCount As Long
    Dim deletedCount As Long
    Dim failedCount As Long
    Dim missingCount As Long

    Dim prompt As String
    Dim response As VbMsgBoxResult

    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\delete extraneous views.txt", True)
    LogPrint logFile, Now & " - Starting delete extraneous views process..."

    ' Connect to Composer and get document
    On Error Resume Next
    Set swApp = New DSComposer.Application
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to create DSComposer app: " & Err.Description
        MsgBox "Failed to connect to SolidWorks Composer: " & Err.Description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo 0
    LogPrint logFile, Now & " - Connected to Composer"

    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - No active document!"
        MsgBox "No active document in SolidWorks Composer. Please ensure the document is open and active (click its window).", vbCritical
        GoTo Cleanup
    End If
    LogPrint logFile, Now & " - Connected to active document"

    ' Gather user-pasted list
    prompt = "Paste view names to delete." & vbCrLf & _
             "Accepted separators: new line, comma, semicolon, or #" & vbCrLf & vbCrLf & _
             "Example:" & vbCrLf & _
             "View 2" & vbCrLf & _
             "View 9"
    rawInput = InputBox(prompt, "Skycode - Delete Extraneous Views")

    If Len(Trim$(rawInput)) = 0 Then
        LogPrint logFile, Now & " - No input provided. Operation canceled by user."
        MsgBox "No values were provided. Nothing was deleted.", vbInformation
        GoTo Cleanup
    End If

    Set targetMap = BuildTargetNameMap(rawInput, logFile)
    requestedCount = targetMap.Count
    If requestedCount = 0 Then
        LogPrint logFile, Now & " - Input parsed to zero usable names."
        MsgBox "No valid view names were found in the pasted list.", vbExclamation
        GoTo Cleanup
    End If
    LogPrint logFile, Now & " - Parsed " & requestedCount & " unique target names from user input"

    ' Get all views from the active document
    LogPrint logFile, Now & " - Retrieving views"
    On Error Resume Next
    viewSelection = swDoc.GetAllViews
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to retrieve views: " & Err.Description
        MsgBox "Failed to retrieve views: " & Err.Description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo 0

    If Not TryParseViewSelection(viewSelection, viewNames, logFile) Then
        LogPrint logFile, Now & " - No views found in document"
        MsgBox "No views found in the active document.", vbInformation
        GoTo Cleanup
    End If

    Set foundMap = CreateObject("Scripting.Dictionary")
    foundMap.CompareMode = vbTextCompare

    LogPrint logFile, Now & " - Matching requested names against document views"
    matchedCount = 0
    For i = 0 To UBound(viewNames)
        key = Trim$(viewNames(i))
        If Len(key) > 0 Then
            If targetMap.Exists(key) Then
                matchedCount = matchedCount + 1
                If Not foundMap.Exists(key) Then foundMap.Add key, True
            End If
        End If
    Next i

    prompt = "Requested names: " & requestedCount & vbCrLf & _
             "Matching views found: " & matchedCount & vbCrLf & vbCrLf & _
             "Proceed to delete matching views?"
    response = MsgBox(prompt, vbQuestion + vbYesNo, "Confirm View Deletion")
    If response <> vbYes Then
        LogPrint logFile, Now & " - User canceled deletion after preview"
        MsgBox "Deletion canceled. No views were deleted.", vbInformation
        GoTo Cleanup
    End If

    ' Delete matching views
    LogPrint logFile, Now & " - Deleting matching views"
    deletedCount = 0
    failedCount = 0

    For i = 0 To UBound(viewNames)
        key = Trim$(viewNames(i))
        If Len(key) > 0 Then
            If targetMap.Exists(key) Then
                On Error Resume Next
                If swDoc.DeleteView(key) Then
                    deletedCount = deletedCount + 1
                    LogPrint logFile, Now & " - Deleted view: '" & key & "'"
                Else
                    failedCount = failedCount + 1
                    LogPrint logFile, Now & " - Failed to delete view '" & key & "': " & Err.Description
                End If
                On Error GoTo 0
            End If
        End If
    Next i

    missingCount = requestedCount - foundMap.Count

    LogPrint logFile, Now & " - Deletion complete"
    LogPrint logFile, Now & " - Summary: requested=" & requestedCount & _
                      ", matched=" & matchedCount & _
                      ", deleted=" & deletedCount & _
                      ", failed=" & failedCount & _
                      ", missing=" & missingCount

    MsgBox "Delete extraneous views complete." & vbCrLf & vbCrLf & _
           "Requested names: " & requestedCount & vbCrLf & _
           "Matching views: " & matchedCount & vbCrLf & _
           "Deleted: " & deletedCount & vbCrLf & _
           "Failed deletes: " & failedCount & vbCrLf & _
           "Not found in document: " & missingCount, vbInformation

Cleanup:
    LogPrint logFile, Now & " - Cleaning up resources..."

    If Not logFile Is Nothing Then
        On Error Resume Next
        logFile.Close
        On Error GoTo 0
    End If

    Set foundMap = Nothing
    Set targetMap = Nothing
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

Private Function BuildTargetNameMap(ByVal rawInput As String, ByVal logFile As Object) As Object
    Dim normalized As String
    Dim tokens() As String
    Dim i As Long
    Dim token As String
    Dim map As Object

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    normalized = rawInput
    normalized = Replace(normalized, vbCrLf, "#")
    normalized = Replace(normalized, vbCr, "#")
    normalized = Replace(normalized, vbLf, "#")
    normalized = Replace(normalized, ",", "#")
    normalized = Replace(normalized, ";", "#")
    normalized = Replace(normalized, vbTab, "#")

    tokens = Split(normalized, "#")
    For i = 0 To UBound(tokens)
        token = Trim$(tokens(i))
        If Len(token) > 0 Then
            If Not map.Exists(token) Then
                map.Add token, True
                LogPrint logFile, Now & " - Target name parsed: '" & token & "'"
            End If
        End If
    Next i

    Set BuildTargetNameMap = map
End Function

Private Function TryParseViewSelection(ByVal viewSelection As String, ByRef viewNames() As String, ByVal logFile As Object) As Boolean
    Dim payload As String
    Dim startPos As Long
    Dim endPos As Long

    payload = viewSelection
    LogPrint logFile, Now & " - Raw view selection (length " & Len(payload) & "): '" & payload & "'"

    startPos = InStr(payload, ">")
    endPos = InStrRev(payload, "<")
    If startPos > 0 And endPos > startPos Then
        payload = Mid$(payload, startPos + 1, endPos - startPos - 1)
    End If

    If Left$(payload, 1) = "#" Then payload = Mid$(payload, 2)
    If Len(payload) > 0 Then
        If Right$(payload, 1) = "#" Then payload = Left$(payload, Len(payload) - 1)
    End If

    If Len(Trim$(payload)) = 0 Then
        TryParseViewSelection = False
        Exit Function
    End If

    viewNames = Split(payload, "#")
    TryParseViewSelection = (UBound(viewNames) >= 0)
End Function

Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub
