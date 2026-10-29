Option Compare Database
Option Explicit

Sub CreateAssemblyViews()
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
    LogPrint logFile, Now & " - Starting view creation..."

    ' Connect to Composer and get document
    Set swApp = New DSComposer.Application
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - No active document!"
        Exit Sub
    End If

    ' Get all actors
    swDoc.AssySelectionMode = 1
    Dim actorSelection As String
    actorSelection = swDoc.GetAllActors(1)
    actors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    
    ' *** Step 1: Identify Top Sub-Assy in Tree (acts as root assy if file has more than one imported Sldwrks config & sub-assy has same base part no.) ***
    Set uniqueAssemblies = New Collection
    If UBound(actors) >= 0 Then
        Dim firstActor As String
        firstActor = Trim(actors(0))
        If Len(firstActor) > 0 Then
            rootID = Split(firstActor, ".")(0) ' e.g., "ADefault"
            LogPrint logFile, Now & " - Step 1: Identifying Top Sub-Assy in Tree... identified: " & rootID
            On Error Resume Next
            uniqueAssemblies.Add rootID, rootID
            On Error GoTo 0
        End If
    End If

    ' *** Step 2: Collect Unique Sub-Assemblies ***
    LogPrint logFile, Now & " - Step 2: Collecting Unique Sub-Assemblies... "
    For i = 0 To UBound(actors)
        actorPath = Trim(actors(i))
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            Dim assemblyPath As String
            assemblyPath = ""
            ' Build assembly paths, stop before the part (last segment)
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

    ' *** Step 3: Top Sub-Assy in Tree ***
    LogPrint logFile, Now & " - Step 3: Processing Top Sub-Assy in Tree: " & rootID
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
        End If
    Else
        LogPrint logFile, Now & " - Failed to select top sub-assy: " & rootID
    End If

    ' *** Step 4: Process Sub-Assemblies ***
    LogPrint logFile, Now & " - Step 4: Processing Sub-Assemblies: "
    Dim assemblyID As Variant
    For Each assemblyID In uniqueAssemblies
        If assemblyID <> rootID Then ' Skip the root, already processed
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
                End If
            Else
                LogPrint logFile, Now & " - Failed to select sub-assembly: " & assemblyID
            End If
        End If
    Next assemblyID

    ' *** Step 5: Process Parts (with don't duplicate routine) ***
LogPrint logFile, Now & " - Step 5: Processing Parts (non-duplicates only) "
Dim processedParts As New Collection ' Tracks unique Meta.PartNo values to avoid duplicates
For i = 0 To UBound(actors)
    actorPath = Trim(actors(i))
    If Len(actorPath) > 0 Then
        segments = Split(actorPath, ".")
        If UBound(segments) >= 0 And Left(segments(UBound(segments)), 1) = "P" Then
            ' Get and clean Meta.PartNo for the part
            Dim partPartNo As String
            partPartNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & actorPath & "</CLitSelection>", "Meta.PartNo")
            Dim cleanPartPartNo As String
            cleanPartPartNo = StripMetaPartNo(partPartNo)
            
            ' Use Meta.PartNo for duplicate check, fallback to actorPath if empty
            Dim partKey As String
            If Len(cleanPartPartNo) > 0 Then
                partKey = cleanPartPartNo
            Else
                partKey = actorPath ' Use full path if no Meta.PartNo
            End If
            
            On Error Resume Next
            processedParts.Add partKey, partKey
            If Err.Number = 0 Then
                ' New unique part, process it
                LogPrint logFile, Now & " - Processing part: " & actorPath & " (Meta.PartNo: '" & cleanPartPartNo & "')"
                If swDoc.SelectOnFilter("<CLitSelection Name="""">" & actorPath & "</CLitSelection>", False) Then
                    swDoc.SetVisibility "", 0, 2
                    swDoc.SetVisibility "<CLitSelection Name="""">" & actorPath & "</CLitSelection>", 1, 2
                    swDoc.ZoomFitAll
                    Dim partViewName As String
                    If Len(cleanPartPartNo) > 0 Then
                        partViewName = cleanPartPartNo
                    Else
                        partViewName = "Part_" & Replace(actorPath, ".", "_")
                    End If
                    If swDoc.CreateView(partViewName, -1) Then
                        LogPrint logFile, Now & " - Created view: " & partViewName & " for " & actorPath
                    End If
                End If
            Else
                ' Duplicate part, skip it
                LogPrint logFile, Now & " - Skipping duplicate part: " & actorPath & " (Meta.PartNo: '" & cleanPartPartNo & "')"
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
Next i

    LogPrint logFile, Now & " - CreateViews for Sub Assemblies & Parts completed."
    logFile.Close
End Sub

' Helper function to strip XML tags and extract the Meta.PartNo value
Function StripMetaPartNo(partNo As String) As String
    Dim startPos As Long, endPos As Long
    ' Look for Value=" in the string
    startPos = InStr(partNo, "Value=""")
    If startPos > 0 Then
        ' Skip past Value=" to the actual value
        startPos = startPos + Len("Value=""")
        ' Find the closing quote
        endPos = InStr(startPos, partNo, """")
        If endPos > startPos Then
            ' Extract the value between the quotes and trim spaces
            StripMetaPartNo = Trim(Mid(partNo, startPos, endPos - startPos))
            Exit Function
        End If
    End If
    ' Fallback: remove XML tags if Value=" is not found
    Dim temp As String
    temp = Replace(partNo, "<Meta.PartNo>", "")
    temp = Replace(temp, "</Meta.PartNo>", "")
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
