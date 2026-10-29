Option Compare Database
Option Explicit

' ============================================================
' Skycode - Rename Views in Numerical Order
' For actors in part selection mode:
'   - Rename all views in the order they are returned by GetAllViews, starting from "View 1", "View 2", etc.
' ============================================================

Sub RenameViewsSequentially()
    ' Declare variables
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim logFile As Object
    Dim fs As Object
    Dim viewNames() As String
    Dim viewSelection As String
    Dim viewCount As Long
    Dim newViewName As String
    Dim i As Long
    Dim startPos As Long
    Dim endPos As Long
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\rename views in numerical order.txt", True)
    LogPrint logFile, Now & " - Starting view renaming process..."
    
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
    
    ' Get all views
    LogPrint logFile, Now & " - Retrieving views"
    On Error Resume Next
    viewSelection = swDoc.GetAllViews
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to retrieve views: " & Err.Description
        MsgBox "Failed to retrieve views: " & Err.Description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo 0
    LogPrint logFile, Now & " - Raw view selection (length " & Len(viewSelection) & "): '" & viewSelection & "'"
    
    ' Process views
    If Len(viewSelection) > 0 Then
        ' Remove CLitSelection wrapper tags if present, then trim delimiters
        startPos = InStr(viewSelection, ">")
        endPos = InStrRev(viewSelection, "<")
        If startPos > 0 And endPos > startPos Then
            viewSelection = Mid(viewSelection, startPos + 1, endPos - startPos - 1)
        End If

        ' Strip any leading/trailing # before splitting to avoid empty first/last elements
        If Left(viewSelection, 1) = "#" Then viewSelection = Mid(viewSelection, 2)
        If Right(viewSelection, 1) = "#" Then viewSelection = Left(viewSelection, Len(viewSelection) - 1)
        viewNames = Split(viewSelection, "#") ' Assuming # delimiter, as in GetAllActors
        LogPrint logFile, Now & " - Found " & UBound(viewNames) + 1 & " views after split"
        For i = 0 To UBound(viewNames)
            viewNames(i) = Trim(viewNames(i))
            LogPrint logFile, Now & " - View " & i & ": '" & viewNames(i) & "'"
        Next i
    Else
        LogPrint logFile, Now & " - No views found!"
        MsgBox "No views found in the document.", vbCritical
        GoTo Cleanup
    End If
    
    ' Rename views
    LogPrint logFile, Now & " - Renaming views"
    viewCount = 1 ' Start naming from View 1
    For i = 0 To UBound(viewNames)
        If Len(viewNames(i)) > 0 Then
            newViewName = "View " & viewCount
            LogPrint logFile, Now & " - Renaming view '" & viewNames(i) & "' to: '" & newViewName & "'"
            On Error Resume Next
            If swDoc.RenameView(viewNames(i), newViewName) Then
                LogPrint logFile, Now & " - Successfully renamed view to: '" & newViewName & "'"
            Else
                LogPrint logFile, Now & " - Failed to rename view to '" & newViewName & "': " & Err.Description
            End If
            On Error GoTo 0
            viewCount = viewCount + 1
        Else
            LogPrint logFile, Now & " - Skipped empty view name at index " & i
        End If
    Next i
    
    ' Log completion
    LogPrint logFile, Now & " - View renaming complete. Total views renamed: " & (viewCount - 1)



Cleanup:
    ' Cleanup
    LogPrint logFile, Now & " - Cleaning up resources..."
    If Not logFile Is Nothing Then
        On Error Resume Next
        logFile.Close
        On Error GoTo 0
    End If
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

' Helper logging function (from sample code)
Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub
