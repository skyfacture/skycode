Attribute VB_Name = "Test 2 - Accessing All Nested Subs recursively"
Option Compare Database
Option Explicit

Sub CreateAssemblyViews()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actorSelection As String
    Dim actors() As String
    Dim actorCount As Long
    Dim i As Long
    Dim selectResult As Boolean
    Dim processedPartNos As Collection
    Dim fs As Object, logFile As Object
    Dim allActorIDs As Collection
    
    On Error GoTo ErrorHandler
    
    ' Initialize logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim logPath As String
    logPath = "C:\Temp\ComposerDebug.txt"
    If Not fs.FolderExists("C:\Temp") Then
        fs.CreateFolder "C:\Temp"
    End If
    Set logFile = fs.CreateTextFile(logPath, True)
    If logFile Is Nothing Then
        Debug.Print "Failed to create log file at " & logPath
    Else
        LogPrint logFile, "Starting assembly view creation process..."
    End If
    
    ' Connect to Composer
    Set swApp = New DSComposer.Application
    If swApp Is Nothing Then
        LogPrint logFile, "Failed to connect to Composer!"
        Exit Sub
    End If
    LogPrint logFile, "Connected to Composer!"
    
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, "No active document found - ensure an .smg file is open in Composer."
        Exit Sub
    End If
    LogPrint logFile, "GetActiveDocument worked!"
    LogPrint logFile, "FileName: " & swDoc.FileName
    
    ' Get all actors
    swDoc.AssySelectionMode = 1
    actorSelection = swDoc.GetAllActors(1)
    If Len(actorSelection) = 0 Then
        LogPrint logFile, "No actors found with Assy.SelectionMode = 1."
        Exit Sub
    End If
    LogPrint logFile, "Raw Assembly Selection: " & Left(actorSelection, 200) & "..."
    
    Dim rawActors() As String
    rawActors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    Dim validCount As Long
    validCount = 0
    For i = 0 To UBound(rawActors)
        If Len(Trim(rawActors(i))) > 0 Then
            validCount = validCount + 1
        End If
    Next i
    
    ReDim actors(validCount - 1)
    Dim j As Long
    j = 0
    For i = 0 To UBound(rawActors)
        If Len(Trim(rawActors(i))) > 0 Then
            actors(j) = rawActors(i)
            j = j + 1
        End If
    Next i
    actorCount = validCount
    LogPrint logFile, "Found " & actorCount & " actors with Assy.SelectionMode = 1."
    
    ' Store all actor IDs for recursion
    Set allActorIDs = New Collection
    For i = 0 To UBound(actors)
        allActorIDs.Add actors(i)
    Next i
    
    ' Initialize processed part numbers
    Set processedPartNos = New Collection
    
    ' Process top-level assemblies
    For i = 0 To actorCount - 1
        Dim cleanActorID As String
        cleanActorID = Replace(actors(i), "<CLitSelection Name="" Ident=""-1"">", "")
        Dim segments() As String
        segments = Split(cleanActorID, ".")
        Dim topLevelID As String
        If UBound(segments) >= 1 Then
            topLevelID = segments(0) & "." & segments(1)
        Else
            topLevelID = cleanActorID
        End If
        LogPrint logFile, "Actor " & i & " Raw ID: " & actors(i)
        LogPrint logFile, "Actor " & i & " Clean ID: " & cleanActorID
        LogPrint logFile, "Actor " & i & " Sub-Assembly ID: " & topLevelID
        
        Dim partNo As String
        partNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", "Meta.PartNo")
        Dim cleanPartNo As String
        cleanPartNo = StripMetaPartNo(partNo)
        LogPrint logFile, "Meta.PartNo (Sub-Assembly ID): '" & cleanPartNo & "'"
        
        Dim actorName As String
        actorName = swDoc.GetPropertyValue("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", "Actor.Name")
        Dim cleanActorName As String
        cleanActorName = StripMetaPartNo(actorName)
        LogPrint logFile, "Actor.Name (Sub-Assembly ID): '" & cleanActorName & "'"
        
        ' Check if it's an assembly based on existing filter
        Dim isAssembly As Boolean
        isAssembly = (cleanPartNo = "7394 1000 38" Or _
                      cleanPartNo = "BG 003 699 42" Or _
                      cleanPartNo = "BG 003 942 73" Or _
                      cleanPartNo = "7372 2702 01" Or _
                      cleanActorName = "7394 1000 38" Or _
                      cleanActorName = "BG 003 699 42" Or _
                      cleanActorName = "BG 003 942 73" Or _
                      cleanActorName = "7372 2702 01" Or _
                      topLevelID = "ADefault" Or _
                      topLevelID = "ABG 003 699 42" Or _
                      topLevelID = "ABG 003 942 73" Or _
                      topLevelID = "A7372 2702 01")
        
        If Not isAssembly Then
            LogPrint logFile, "Skipping non-assembly actor " & i & ": " & cleanActorID & " (PartNo: '" & cleanPartNo & "')"
        Else
            ' Check for duplicates and process recursively
            On Error Resume Next
            processedPartNos.Add cleanPartNo, cleanPartNo
            If Err.Number = 0 Then
                LogPrint logFile, "Processing assembly " & i & ": " & cleanActorID & " (Meta.PartNo: '" & cleanPartNo & "')"
                ProcessAssembly topLevelID, swDoc, allActorIDs, processedPartNos, logFile
            Else
                LogPrint logFile, "Skipping duplicate assembly " & i & ": " & cleanActorID & " (Meta.PartNo: '" & cleanPartNo & "')"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next i
    
    LogPrint logFile, "Assembly view creation process completed."
    If Not logFile Is Nothing Then
        logFile.Close
    End If
    Set logFile = Nothing
    Set fs = Nothing
    Exit Sub

ErrorHandler:
    If Not logFile Is Nothing Then
        LogPrint logFile, "Error occurred: " & Err.Number & " - " & Err.Description
        logFile.Close
    Else
        Debug.Print "Error occurred before log file creation: " & Err.Number & " - " & Err.Description
    End If
    Set logFile = Nothing
    Set fs = Nothing
    On Error GoTo 0
End Sub

' Recursive subroutine to process an assembly and its sub-assemblies
Sub ProcessAssembly(parentID As String, swDoc As DSComposer.Document, allActorIDs As Collection, processedPartNos As Collection, logFile As Object)
    Dim selectResult As Boolean
    
    ' Select the assembly
    selectResult = swDoc.SelectOnFilter("<CLitSelection Name="""">" & parentID & "</CLitSelection>", False)
    If Not selectResult Then
        LogPrint logFile, "Failed to select assembly: " & parentID
        Exit Sub
    End If
    LogPrint logFile, "Selected assembly: " & parentID
    
    ' Set visibility: hide all, show only this assembly and its children
    swDoc.SetVisibility "", 0, 2
    swDoc.SetVisibility "<CLitSelection Name="""">" & parentID & "</CLitSelection>", 1, 2
    
    ' Zoom to fit
    swDoc.ZoomFitAll
    LogPrint logFile, "Zoomed to assembly: " & parentID
    
    ' Get part number for view naming
    Dim viewName As String
    viewName = GetPartNo(parentID, swDoc)
    If Len(viewName) = 0 Then
        viewName = "View_" & Replace(parentID, ".", "_") ' Fallback name
    End If
    
    ' Create view
    If swDoc.CreateView(viewName, -1) Then ' -1 to append at end
        LogPrint logFile, "Created view '" & viewName & "' for " & parentID
    Else
        LogPrint logFile, "Failed to create view for " & parentID
        Exit Sub
    End If
    
    ' Find and process immediate sub-assemblies
    Dim subAssys As Collection
    Set subAssys = GetImmediateSubAssemblies(parentID, allActorIDs)
    Dim subAssy As Variant
    For Each subAssy In subAssys
        ProcessAssembly CStr(subAssy), swDoc, allActorIDs, processedPartNos, logFile
    Next subAssy
End Sub

' Function to get immediate sub-assemblies under a parent ID
Function GetImmediateSubAssemblies(parentID As String, allActorIDs As Collection) As Collection
    Dim subAssys As New Collection
    Dim parentSegments() As String
    parentSegments = Split(parentID, ".")
    Dim parentLevel As Integer
    parentLevel = UBound(parentSegments) + 1
    
    Dim actorID As Variant
    For Each actorID In allActorIDs
        Dim segments() As String
        segments = Split(CStr(actorID), ".")
        If UBound(segments) >= parentLevel Then
            ' Check if this actor is a direct child of parentID
            Dim match As Boolean
            match = True
            Dim k As Integer
            For k = 0 To parentLevel - 1
                If segments(k) <> parentSegments(k) Then
                    match = False
                    Exit For
                End If
            Next k
            If match Then
                Dim nextSegment As String
                nextSegment = segments(parentLevel)
                If Left(nextSegment, 1) <> "P" Then ' Sub-assembly, not a part
                    Dim subAssyID As String
                    subAssyID = parentID & "." & nextSegment
                    On Error Resume Next
                    subAssys.Add subAssyID, subAssyID ' Avoid duplicates
                    On Error GoTo 0
                End If
            End If
        End If
    Next actorID
    Set GetImmediateSubAssemblies = subAssys
End Function

' Function to get cleaned part number for view naming
Function GetPartNo(actorID As String, swDoc As DSComposer.Document) As String
    Dim partNo As String
    partNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & actorID & "</CLitSelection>", "Meta.PartNo")
    If Len(partNo) > 0 Then
        GetPartNo = StripMetaPartNo(partNo)
    Else
        Dim actorName As String
        actorName = swDoc.GetPropertyValue("<CLitSelection Name="""">" & actorID & "</CLitSelection>", "Actor.Name")
        GetPartNo = StripMetaPartNo(actorName)
    End If
End Function

' Existing logging subroutine (unchanged)
Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    If Not logFile Is Nothing Then
        logFile.WriteLine Now & " - " & message
    Else
        Debug.Print "Log file not initialized: " & message
    End If
End Sub

' Existing part number stripping function (unchanged)
Function StripMetaPartNo(partNo As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim strippedValue As String
    
    startPos = InStr(partNo, "Value=""") + 7
    endPos = InStr(startPos, partNo, """")
    If startPos > 7 And endPos > startPos Then
        strippedValue = Mid(partNo, startPos, endPos - startPos)
        StripMetaPartNo = strippedValue
    Else
        strippedValue = Replace(Replace(partNo, "<Meta.PartNo>", ""), "</Meta.PartNo>", "")
        If Len(strippedValue) > 0 Then
            StripMetaPartNo = strippedValue
        Else
            StripMetaPartNo = ""
        End If
    End If
End Function

