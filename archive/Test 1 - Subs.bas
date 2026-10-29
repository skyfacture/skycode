Attribute VB_Name = "Test 1 - Subs"
Option Explicit

Sub QueryManuallySelectedObject()
    ' Variables for Composer and selection
    Dim swApp As Object ' DSComposer.Application
    Dim swDoc As Object ' DSComposer.Document
    Dim selectedActors As String
    Dim actorArray() As String
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Note: All messages and errors are output to the Immediate Window (Ctrl+G in VBA editor)
    
    ' ### Step 1: Connect to Composer ###
    Set swApp = CreateObject("DSComposer.Application")
    If swApp Is Nothing Then
        Debug.Print "Cannot connect to Composer!"
        Exit Sub
    End If
    Debug.Print "Connected to Composer."
    
    ' ### Step 2: Get the active document ###
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        Debug.Print "No active document found. Open a file!"
        Exit Sub
    End If
    Debug.Print "Active document: " & swDoc.FileName
    
    ' ### Step 3: Get the manually selected object(s) ###
    ' Use 1 to get only selected actors
    selectedActors = swDoc.GetAllActors(1)
    If Len(selectedActors) = 0 Then
        Debug.Print "Nothing is selected. Select an object!"
        Exit Sub
    End If
    Debug.Print "Raw selection: " & Left(selectedActors, 200) & "..."
    
    ' ### Step 4: Parse the selection to extract IDs ###
    ' Remove XML tags and split by "#" to get IDs
    actorArray = Split(Replace(Replace(selectedActors, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    
    ' ### Step 5: Display the IDs in the Immediate Window ###
    If UBound(actorArray) >= 0 Then
        Debug.Print "Selected object ID(s): " & Join(actorArray, ", ")
    Else
        Debug.Print "Failed to parse the selection."
    End If
    
    Exit Sub
    
' ### Error Handling ###
ErrorHandler:
    Debug.Print "Error: " & Err.Description
End Sub
