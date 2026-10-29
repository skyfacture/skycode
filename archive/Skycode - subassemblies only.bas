Option Compare Database
Option Explicit

Sub CreateSubAssemblyViews()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actors() As String
    Dim i As Long, j As Long
    Dim logFile As Object, fs As Object
    Dim uniqueAssemblies As Collection
    Dim rootID As String
    Dim actorPath As String
    Dim segments() As String
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\Temp\ComposerDebug.txt", True)
    LogPrint logFile, Now & " - Starting sub-assembly view creation..."

    ' Connect to Composer and get document
    Set swApp = New DSComposer.Application
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - No active document!"
        Exit Sub
    End If

    ' Get all actors for assemblies
    swDoc.AssySelectionMode = 1
    Dim actorSelection As String
    actorSelection = swDoc.GetAllActors(1)
    LogPrint logFile, Now & " - Raw actorSelection for sub-assemblies (length " & Len(actorSelection) & "): '" & actorSelection & "'"
    actors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    If UBound(actors) >= 0 Then
        LogPrint logFile, Now & " - Found " & UBound(actors) + 1 & " actors for sub-assemblies"
        For i = 0 To UBound(actors)
            actors(i) = Trim(Replace(actors(i), "<CLitSelection Name="""" Ident=""-1"">", ""))
            LogPrint logFile, Now & " - Sub-assembly actor " & i & ": '" & actors(i) & "'"
        Next i
    Else
        LogPrint logFile, Now & " - No sub-assembly actors found!"
        GoTo Cleanup
    End If
    
    ' Step 1: Identify Top Sub-Assembly in Tree
    Set uniqueAssemblies = New Collection
    If UBound(actors) >= 0 Then
        Dim firstActor As String
        firstActor = Trim(actors(0))
        If Len(firstActor) > 0 Then
            segments = Split(firstActor, ".")
            If UBound(segments) >= 0 Then
                rootID = segments(0)
                LogPrint logFile, Now & " - Step 1: Identified Top Sub-Assembly: " & rootID
                On Error Resume Next
                uniqueAssemblies.Add rootID, rootID
                On Error GoTo 0
            End If
        End If
    End If

    ' Step 2: Collect Unique Sub-Assemblies
    LogPrint logFile, Now & " - Step 2: Collecting Unique Sub-Assemblies..."
    For i = 0 To UBound(actors)
        actorPath = Trim(actors(i))
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            Dim assemblyPath As String
            assemblyPath = ""
            For j = 0 To UBound(segments) - 1
                If assemblyPath = "" Then
                    assemblyPath = segments(j)
                Else
                    assemblyPath = assemblyPath & "." & segments(j)
                End If
                On Error Resume Next
                uniqueAssemblies.Add assemblyPath, assemblyPath
                On Error GoTo 0
            Next j
        End If
    Next i

    ' Step 3: Process Top Sub-Assembly
    LogPrint logFile, Now & " - Step 3: Processing Top Sub-Assembly: " & rootID
    If Len(rootID) > 0 Then
        If swDoc.SelectOnFilter("<CLitSelection Name="""">" & rootID & "</CLitSelection>", False) Then
            swDoc.SetVisibility "", 0, 2
            swDoc.SetVisibility "<CLitSelection Name="""">" & rootID & "</CLitSelection>", 1, 2
            swDoc.ZoomFitAll
            Dim rootPartNo As String
            rootPartNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & rootID & "</CLitSelection>", "Meta.PartNo")
            Dim cleanRootPartNo As String
            cleanRootPartNo = StripMetaPartNo(rootPartNo)
            Dim rootViewName As String
            If Len(cleanRootPartNo) > 0 Then
                rootViewName = cleanRootPartNo
            Else
                rootViewName = "Root_" & Replace(rootID, ".", "_")
                LogPrint logFile, Now & " - Root has no Meta.PartNo, using: " & rootViewName
            End If
            If swDoc.CreateView(rootViewName, -1) Then
                LogPrint logFile, Now & " - Created view: " & rootViewName & " for " & rootID
            Else
                LogPrint logFile, Now & " - Failed to create view: " & rootViewName & " for " & rootID
            End If
        Else
            LogPrint logFile, Now & " - Failed to select top sub-assembly: " & rootID
        End If
    Else
        LogPrint logFile, Now & " - No valid rootID!"
    End If

    ' Step 4: Process Sub-Assemblies
    LogPrint logFile, Now & " - Step 4: Processing Sub-Assemblies..."
    Dim assemblyID As Variant
    For Each assemblyID In uniqueAssemblies
        If assemblyID <> rootID Then
            LogPrint logFile, Now & " - Processing sub-assembly: " & assemblyID
            If swDoc.SelectOnFilter("<CLitSelection Name="""">" & assemblyID & "</CLitSelection>", False) Then
                swDoc.SetVisibility "", 0, 2
                swDoc.SetVisibility "<CLitSelection Name="""">" & assemblyID & "</CLitSelection>", 1, 2
                swDoc.ZoomFitAll
                Dim subPartNo As String
                subPartNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & assemblyID & "</CLitSelection>", "Meta.PartNo")
                Dim cleanSubPartNo As String
                cleanSubPartNo = StripMetaPartNo(subPartNo)
                Dim subViewName As String
                If Len(cleanSubPartNo) > 0 Then
                    subViewName = cleanSubPartNo
                Else
                    subViewName = "SubAssy_" & Replace(assemblyID, ".", "_")
                End If
                If swDoc.CreateView(subViewName, -1) Then
                    LogPrint logFile, Now & " - Created view: " & subViewName & " for " & assemblyID
                Else
                    LogPrint logFile, Now & " - Failed to create view: " & subViewName & " for " & assemblyID
                End If
            Else
                LogPrint logFile, Now & " - Failed to select sub-assembly: " & assemblyID
            End If
        End If
    Next assemblyID

Cleanup:
    LogPrint logFile, Now & " - Sub-Assembly View Creation Completed."
    logFile.Close
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

' Helper function to strip XML tags and extract the Meta.PartNo value
Function StripMetaPartNo(partNo As String) As String
    Dim startPos As Long, endPos As Long
    Dim result As String

     startPos = InStr(partNo, "Value=""")
    If startPos > 0 Then
        startPos = startPos + Len("Value=""")
        endPos = InStr(startPos, partNo, """")
        If endPos > startPos Then
            result = Mid(partNo, startPos, endPos - startPos)
            ' Remove all control characters (CR, LF, Tab, etc.)
            result = Replace(result, Chr(13), "")   ' Remove CR character
            result = Replace(result, Chr(10), "")   ' Remove LF character
            result = Replace(result, Chr(9), "")    ' Remove Tab character
            result = Replace(result, vbCr, "")      ' Additional CR removal
            result = Replace(result, vbLf, "")      ' Additional LF removal
            result = Replace(result, vbTab, "")     ' Additional Tab removal
            result = Replace(result, vbCrLf, "")    ' Remove CRLF combo
            ' Remove HTML/XML entities for control characters
            result = Replace(result, "&#13;", "")   ' Remove &#13; entity
            result = Replace(result, "&#10;", "")   ' Remove &#10; entity
            result = Replace(result, "&#9;", "")    ' Remove &#9; entity
            result = Replace(result, "&amp;", "&")  ' Decode &amp;
            StripMetaPartNo = Trim(result)
            Exit Function
        End If
    End If
    
    Dim temp As String
    temp = Replace(partNo, "<Meta.PartNo>", "")
    temp = Replace(temp, "</Meta.PartNo>", "")
    temp = Replace(temp, Chr(13), "")   ' Remove CR character
    temp = Replace(temp, Chr(10), "")   ' Remove LF character
    temp = Replace(temp, Chr(9), "")    ' Remove Tab character
    temp = Replace(temp, vbCr, "")      ' Additional CR removal
    temp = Replace(temp, vbLf, "")      ' Additional LF removal
    temp = Replace(temp, vbTab, "")     ' Additional Tab removal
    temp = Replace(temp, vbCrLf, "")    ' Remove CRLF combo
    ' Remove HTML/XML entities for control characters
    temp = Replace(temp, "&#13;", "")   ' Remove &#13; entity
    temp = Replace(temp, "&#10;", "")   ' Remove &#10; entity
    temp = Replace(temp, "&#9;", "")    ' Remove &#9; entity
    temp = Replace(temp, "&amp;", "&")  ' Decode &amp;
    temp = Trim(temp)
    If Len(temp) > 0 Then
        StripMetaPartNo = temp
    Else
        StripMetaPartNo = ""
    End If
End Function

' Helper logging function
Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub
