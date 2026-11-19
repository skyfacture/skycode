Option Compare Database
Option Explicit

' ============================================================
' Skycode - Delete Extraneous Views
' Deletes views from the active Composer document where view names
' match names listed in a text file.
' ============================================================

Sub DeleteExtraneousViews()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim logFile As Object
    Dim fs As Object
    Dim textFile As Object

    Dim rawInput As String
    Dim inputFilePath As String
    Dim targetMap As Object

    Dim targetNames As Variant
    Dim i As Long
    Dim key As String
    Dim deleteSucceeded As Boolean
    Dim deleteErrorNumber As Long
    Dim deleteError As String

    Dim requestedCount As Long
    Dim deletedCount As Long
    Dim notDeletedCount As Long

    Dim prompt As String
    Dim response As VbMsgBoxResult

    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\skycode\logs\delete extraneous views.txt", True)
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

    ' Read the list of view names from a text file
    inputFilePath = PromptForViewNameFile(logFile)
    If Len(inputFilePath) = 0 Then
        LogPrint logFile, Now & " - No input file selected. Operation canceled by user."
        MsgBox "No file was selected. Nothing was deleted.", vbInformation
        GoTo Cleanup
    End If

    On Error Resume Next
    Set textFile = fs.OpenTextFile(inputFilePath, 1, False)
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to open view-name file '" & inputFilePath & "': " & Err.Description
        MsgBox "Failed to open the selected text file: " & Err.Description, vbCritical
        Err.Clear
        On Error GoTo 0
        GoTo Cleanup
    End If
    rawInput = textFile.ReadAll
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to read view-name file '" & inputFilePath & "': " & Err.Description
        MsgBox "Failed to read the selected text file: " & Err.Description, vbCritical
        Err.Clear
        textFile.Close
        Set textFile = Nothing
        On Error GoTo 0
        GoTo Cleanup
    End If
    textFile.Close
    Set textFile = Nothing
    On Error GoTo 0

    Set targetMap = BuildTargetNameMap(rawInput, logFile)
    requestedCount = targetMap.Count
    If requestedCount = 0 Then
        LogPrint logFile, Now & " - Text file contained zero usable names."
        MsgBox "No valid view names were found in the selected text file.", vbExclamation
        GoTo Cleanup
    End If
    LogPrint logFile, Now & " - Parsed " & requestedCount & " unique target names from text file"

    prompt = "Names loaded from file: " & requestedCount & vbCrLf & vbCrLf & _
             "Proceed to attempt deleting each listed view by name?"
    response = MsgBox(prompt, vbQuestion + vbYesNo, "Confirm View Deletion")
    If response <> vbYes Then
        LogPrint logFile, Now & " - User canceled deletion before attempting requested names"
        MsgBox "Deletion canceled. No views were deleted.", vbInformation
        GoTo Cleanup
    End If

    ' Delete each requested view by name
    LogPrint logFile, Now & " - Attempting to delete requested views by name"
    deletedCount = 0
    notDeletedCount = 0
    targetNames = targetMap.Keys

    For i = LBound(targetNames) To UBound(targetNames)
        key = CStr(targetNames(i))
        Err.Clear
        On Error Resume Next
        deleteSucceeded = swDoc.DeleteView(key)
        deleteErrorNumber = Err.Number
        deleteError = Err.Description
        On Error GoTo 0

        If deleteErrorNumber <> 0 Then
            notDeletedCount = notDeletedCount + 1
            LogPrint logFile, Now & " - Error deleting view '" & key & "': " & deleteError
        ElseIf deleteSucceeded Then
            deletedCount = deletedCount + 1
            LogPrint logFile, Now & " - Deleted view: '" & key & "'"
        Else
            notDeletedCount = notDeletedCount + 1
            LogPrint logFile, Now & " - Composer returned False; view not deleted: '" & key & "'"
        End If
    Next i

    LogPrint logFile, Now & " - Deletion complete"
    LogPrint logFile, Now & " - Summary: requested=" & requestedCount & _
                      ", deleted=" & deletedCount & _
                      ", notDeleted=" & notDeletedCount

    MsgBox "Delete extraneous views complete." & vbCrLf & vbCrLf & _
           "Requested names: " & requestedCount & vbCrLf & _
           "Deleted: " & deletedCount & vbCrLf & _
           "Not deleted: " & notDeletedCount, vbInformation

Cleanup:
    LogPrint logFile, Now & " - Cleaning up resources..."

    If Not logFile Is Nothing Then
        On Error Resume Next
        logFile.Close
        On Error GoTo 0
    End If

    Set targetMap = Nothing
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

Private Function PromptForViewNameFile(ByVal logFile As Object) As String
    Dim picker As Object
    Dim selectedPath As String

    On Error Resume Next
    Set picker = Application.FileDialog(3)
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - File picker unavailable: " & Err.Description
        Err.Clear
        selectedPath = InputBox("Enter the full path to the view-name .txt file:", "Skycode - View Name File")
        PromptForViewNameFile = Trim$(selectedPath)
        Exit Function
    End If
    On Error GoTo 0

    With picker
        .Title = "Select View Name Text File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        If .Show = -1 Then
            PromptForViewNameFile = .SelectedItems(1)
        Else
            PromptForViewNameFile = ""
        End If
    End With
End Function

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

Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub
