Option Compare Database
Option Explicit

Private Const VERBOSE_LOG As Boolean = False

' ============================================================
' Skycode - Views [Create Balance Parts Not In DB]
' Creates views only for part numbers in text file provided by the user.
' ============================================================

Sub CreateViewsForPartsNotInDb()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim logFile As Object
    Dim fs As Object

    Dim partNumberFilePath As String
    Dim targetPartMap As Object
    Dim matchedPartMap As Object

    Dim actorSelection As String
    Dim actorsAssy() As String
    Dim actorsPart() As String
    Dim uniqueAssemblies As Object
    Dim actorToPartMap As Object
    Dim preferredActorMap As Object
    Dim i As Long
    Dim actorPath As String
    Dim requestedPart As Variant
    Dim selectionString As String
    Dim canonicalPartNo As String
    Dim viewName As String

    Dim requestedCount As Long
    Dim matchedCount As Long
    Dim matchedActorCount As Long
    Dim createdCount As Long
    Dim failedCount As Long
    Dim missingCount As Long
    Dim processedRootPartNo As String

    Dim prompt As String
    Dim response As VbMsgBoxResult
    Dim step1Status As String

    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\views [create balance parts not in db].txt", True)
    LogPrint logFile, Now & " - Starting create-views-for-pasted-parts process..."

    ' Connect to Composer and get document
    On Error Resume Next
    Set swApp = New DSComposer.Application
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to create DSComposer app: " & Err.description
        MsgBox "Failed to connect to SolidWorks Composer: " & Err.description, vbCritical
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

    ' Gather text file path containing line-separated part numbers
    partNumberFilePath = PromptForPartNumberFile(logFile)
    If Len(partNumberFilePath) = 0 Then
        LogPrint logFile, Now & " - No part-number file selected. Operation canceled by user."
        MsgBox "No part-number text file was selected. Nothing was created.", vbInformation
        GoTo Cleanup
    End If
    LogPrint logFile, Now & " - Selected part-number file: '" & partNumberFilePath & "'"

    Set targetPartMap = BuildTargetNameMapFromTextFile(partNumberFilePath, logFile)
    requestedCount = targetPartMap.Count
    If requestedCount = 0 Then
        LogPrint logFile, Now & " - File parsed to zero usable part numbers."
        MsgBox "No valid part numbers were found in the selected text file.", vbExclamation
        GoTo Cleanup
    End If
    LogPrint logFile, Now & " - Parsed " & requestedCount & " unique target part numbers from file"

    ' Step 1: identify root, match it against requested part numbers, and create+rename root view.
    processedRootPartNo = ProcessRootViewIfRequested(swDoc, targetPartMap, logFile)
    If Len(processedRootPartNo) > 0 Then
        step1Status = "Root view created and renamed to: " & processedRootPartNo
    Else
        step1Status = "Root view skipped (no match, no root part number, or create/rename failed)."
    End If

    If STEP1_ONLY_TEST Then
        LogPrint logFile, Now & " - STEP1_ONLY_TEST=True. Skipping Step 2 actor-cache processing."
        MsgBox "Step 1 test complete." & vbCrLf & vbCrLf & step1Status, vbInformation
        GoTo Cleanup
    End If

    ' Pass 1: assemblies
    swDoc.AssySelectionMode = 1
    LogPrint logFile, Now & " - Retrieving assembly actors (Pass 1, AssySelectionMode=1)"
    On Error Resume Next
    actorSelection = swDoc.GetAllActors(1)
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to retrieve assembly actors: " & Err.description
        MsgBox "Failed to retrieve assembly actors: " & Err.description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo 0

    If Not TryParseSelection(actorSelection, actorsAssy, logFile, "assembly") Then
        LogPrint logFile, Now & " - No assembly actors found in document"
        MsgBox "No assembly actors found in the active document.", vbInformation
        GoTo Cleanup
    End If

    Set uniqueAssemblies = CollectUniqueAssemblyPaths(actorsAssy, logFile)
    SeedRootAssemblyPath swDoc, uniqueAssemblies, logFile

    ' Pass 2: parts
    swDoc.AssySelectionMode = 0
    LogPrint logFile, Now & " - Retrieving part actors (Pass 2, AssySelectionMode=0)"
    On Error Resume Next
    actorSelection = swDoc.GetAllActors(0)
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to retrieve part actors: " & Err.description
        MsgBox "Failed to retrieve part actors: " & Err.description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo 0

    If Not TryParseSelection(actorSelection, actorsPart, logFile, "part") Then
        LogPrint logFile, Now & " - No part actors found in document"
        MsgBox "No part actors found in the active document.", vbInformation
        GoTo Cleanup
    End If

    Set matchedPartMap = CreateObject("Scripting.Dictionary")
    matchedPartMap.CompareMode = vbTextCompare

    Set actorToPartMap = CreateObject("Scripting.Dictionary")
    actorToPartMap.CompareMode = vbTextCompare

    Set preferredActorMap = CreateObject("Scripting.Dictionary")
    preferredActorMap.CompareMode = vbTextCompare

    LogPrint logFile, Now & " - Building actor caches (assembly preferred, part fallback)"
    BuildActorCaches swDoc, uniqueAssemblies, actorsPart, targetPartMap, actorToPartMap, preferredActorMap, matchedPartMap, matchedActorCount, logFile
    matchedCount = matchedPartMap.Count

    prompt = "Requested part numbers (from file): " & requestedCount & vbCrLf & _
             "Matching requested part numbers found: " & matchedCount & vbCrLf & vbCrLf & _
             "Proceed to create views for matching parts?"
    response = MsgBox(prompt, vbQuestion + vbYesNo, "Confirm View Creation")
    If response <> vbYes Then
        LogPrint logFile, Now & " - User canceled creation after preview"
        MsgBox "Creation canceled. No views were created.", vbInformation
        GoTo Cleanup
    End If

    createdCount = 0

    ' Create one view per requested part number using preferred actor map
    LogPrint logFile, Now & " - Creating views for matched unique part numbers"
    failedCount = 0

    For Each requestedPart In targetPartMap.Keys
        canonicalPartNo = CStr(requestedPart)
        If Len(processedRootPartNo) > 0 Then
            If StrComp(canonicalPartNo, processedRootPartNo, vbTextCompare) = 0 Then
                GoTo NextRequestedPart
            End If
        End If

        If preferredActorMap.Exists(canonicalPartNo) Then
            actorPath = CStr(preferredActorMap(canonicalPartNo))
            viewName = canonicalPartNo

            If TrySelectActorPath(swDoc, actorPath, selectionString) Then
                swDoc.SetVisibility "", 0, 2
                swDoc.SetVisibility selectionString, 1, 2
                swDoc.ZoomFitAll

                On Error Resume Next
                If swDoc.CreateView(viewName, -1) Then
                    createdCount = createdCount + 1
                    LogVerbose logFile, Now & " - Created view: '" & viewName & "' for actor: '" & actorPath & "'"
                Else
                    failedCount = failedCount + 1
                    LogPrint logFile, Now & " - Failed to create view '" & viewName & "' for actor: '" & actorPath & "': " & Err.description
                End If
                On Error GoTo 0
            Else
                failedCount = failedCount + 1
                LogPrint logFile, Now & " - Failed to select actor for view creation: '" & actorPath & "'"
            End If
        End If
NextRequestedPart:
    Next requestedPart

    missingCount = requestedCount - matchedPartMap.Count

    LogPrint logFile, Now & " - View creation complete"
    LogPrint logFile, Now & " - Summary: requested=" & requestedCount & _
                      ", matchedRequested=" & matchedCount & _
                      ", matchedActors=" & matchedActorCount & _
                      ", created=" & createdCount & _
                      ", failed=" & failedCount & _
                      ", missing=" & missingCount

    MsgBox "Create balance part views complete." & vbCrLf & vbCrLf & _
           "Requested part numbers: " & requestedCount & vbCrLf & _
            "Matching requested part numbers: " & matchedCount & vbCrLf & _
           "Views created: " & createdCount & vbCrLf & _
           "Failed creates: " & failedCount & vbCrLf & _
           "Not found in model: " & missingCount, vbInformation

Cleanup:
    LogPrint logFile, Now & " - Cleaning up resources..."

    If Not logFile Is Nothing Then
        On Error Resume Next
        logFile.Close
        On Error GoTo 0
    End If

    Set preferredActorMap = Nothing
    Set actorToPartMap = Nothing
    Set uniqueAssemblies = Nothing
    Set matchedPartMap = Nothing
    Set targetPartMap = Nothing
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

Private Function ProcessRootViewIfRequested( _
    ByVal swDoc As DSComposer.Document, _
    ByVal targetPartMap As Object, _
    ByVal logFile As Object) As String

    Dim actorSelection As String
    Dim actors() As String
    Dim actorCount As Long
    Dim i As Long
    Dim selectResult As Boolean
    Dim processedPartNos As Collection
    Dim rawActors() As String
    Dim validCount As Long
    Dim j As Long
    Dim cleanActorID As String
    Dim topLevelID As String
    Dim partNo As String
    Dim cleanPartNo As String
    Dim actorName As String
    Dim cleanActorName As String
    Dim isAssembly As Boolean
    Dim initialViewName As String
    Dim newViewName As String

    ProcessRootViewIfRequested = ""

    LogPrint logFile, Now & " - Starting root assembly identification..."

    swDoc.AssySelectionMode = 1
    actorSelection = swDoc.GetAllActors(1)
    If Len(actorSelection) = 0 Then
        LogPrint logFile, Now & " - No actors found with Assy.SelectionMode = 1."
        Exit Function
    End If

    rawActors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    validCount = 0
    For i = 0 To UBound(rawActors)
        If Len(Trim$(rawActors(i))) > 0 Then
            validCount = validCount + 1
        End If
    Next i

    If validCount = 0 Then
        LogPrint logFile, Now & " - No usable actors found in assembly selection."
        Exit Function
    End If

    ReDim actors(validCount - 1)
    j = 0
    For i = 0 To UBound(rawActors)
        If Len(Trim$(rawActors(i))) > 0 Then
            actors(j) = rawActors(i)
            j = j + 1
        End If
    Next i
    actorCount = validCount
    LogPrint logFile, Now & " - Found " & actorCount & " actors with Assy.SelectionMode = 1."

    Set processedPartNos = New Collection

    For i = 0 To actorCount - 1
        cleanActorID = Replace(actors(i), "<CLitSelection Name="" Ident=""-1"">", "")
        topLevelID = Left(cleanActorID, InStr(cleanActorID & ".", ".") - 1)
        topLevelID = Replace(topLevelID, "<CLitSelection Name="" Ident=""-1"">", "")
        LogPrint logFile, Now & " - Actor " & i & " Raw ID: " & actors(i)
        LogPrint logFile, Now & " - Actor " & i & " Clean ID: " & cleanActorID
        LogPrint logFile, Now & " - Actor " & i & " Top-Level ID: " & topLevelID

        partNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", "Meta.PartNo")
        cleanPartNo = StripMetaPartNo(partNo)
        LogPrint logFile, Now & " - Meta.PartNo (Top-Level ID): '" & cleanPartNo & "'"

        actorName = swDoc.GetPropertyValue("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", "Actor.Name")
        cleanActorName = StripMetaPartNo(actorName)
        LogPrint logFile, Now & " - Actor.Name (Top-Level ID): '" & cleanActorName & "'"

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
            LogPrint logFile, Now & " - Skipping non-assembly actor " & i & ": " & cleanActorID & " (PartNo: '" & cleanPartNo & "')"
            GoTo NextActor
        End If

        If Not targetPartMap.Exists(cleanPartNo) Then
            LogPrint logFile, Now & " - Skipping assembly " & i & " because it is not in the request file: '" & cleanPartNo & "'"
            GoTo NextActor
        End If

        On Error Resume Next
        processedPartNos.Add cleanPartNo, cleanPartNo
        If Err.Number = 0 Then
            LogPrint logFile, Now & " - Processing assembly " & i & ": " & cleanActorID & " (Meta.PartNo: '" & cleanPartNo & "')"
        Else
            LogPrint logFile, Now & " - Skipping duplicate assembly " & i & ": " & cleanActorID & " (Meta.PartNo: '" & cleanPartNo & "')"
            Err.Clear
            GoTo NextActor
        End If
        On Error GoTo 0

        selectResult = swDoc.SelectOnFilter("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", False)
        If selectResult Then
            LogPrint logFile, Now & " - Selected assembly " & i & " successfully with Top-Level ID: " & topLevelID
        Else
            LogPrint logFile, Now & " - Failed to select assembly " & i & " with SelectOnFilter on Top-Level ID: " & topLevelID
            processedPartNos.Remove cleanPartNo
            GoTo NextActor
        End If

        swDoc.SetVisibility "", 0, 2
        swDoc.SetVisibility "<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", 1, 2

        swDoc.ZoomFitAll
        If Err.Number <> 0 Then
            LogPrint logFile, Now & " - Zoom failed for assembly " & i & ": " & Err.description
            Err.Clear
            processedPartNos.Remove cleanPartNo
            GoTo NextActor
        End If
        LogPrint logFile, Now & " - Zoomed to assembly " & i & "."

        initialViewName = "View " & (i + 1)
        If swDoc.CreateView(initialViewName, i) Then
            LogPrint logFile, Now & " - Created view '" & initialViewName & "' at position " & i & "."

            If Len(cleanPartNo) > 0 Then
                newViewName = cleanPartNo
                If swDoc.RenameView(initialViewName, newViewName) Then
                    LogPrint logFile, Now & " - Renamed '" & initialViewName & "' to '" & newViewName & "'."
                Else
                    LogPrint logFile, Now & " - Failed to rename '" & initialViewName & "' to '" & newViewName & "'."
                    processedPartNos.Remove cleanPartNo
                End If
            Else
                LogPrint logFile, Now & " - No valid Meta.PartNo for assembly " & i & "; keeping '" & initialViewName & "'."
            End If
        Else
            LogPrint logFile, Now & " - Failed to create view '" & initialViewName & "'."
            processedPartNos.Remove cleanPartNo
        End If

        ProcessRootViewIfRequested = cleanPartNo

NextActor:
    Next i

End Function

Private Sub SeedRootAssemblyPath(ByVal swDoc As DSComposer.Document, ByVal uniqueAssemblies As Object, ByVal logFile As Object)
    Dim actorSelection As String
    Dim actors() As String
    Dim rootID As String
    Dim parentSelection As String
    Dim cleanedSelection As String

    LogPrint logFile, Now & " - Resolving root assembly path for normal matching"
    actorSelection = swDoc.GetAllActorsByType("CLitKAssyRoot")
    If Len(actorSelection) = 0 Then Exit Sub

    cleanedSelection = Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", "")
    actors = Split(cleanedSelection, "#")
    If UBound(actors) < 0 Then Exit Sub

    rootID = Trim$(actors(0))
    If Len(rootID) = 0 Then Exit Sub

    parentSelection = swDoc.GetAssyParents(rootID)
    If Len(parentSelection) > 0 And parentSelection <> "<CLitSelection Name="" Ident=""-1""></CLitSelection>" Then
        rootID = parentSelection
    End If

    If Not uniqueAssemblies.Exists(rootID) Then
        uniqueAssemblies.Add rootID, True
        LogPrint logFile, Now & " - Seeded root assembly path for normal matching: '" & rootID & "'"
    End If
End Sub

Private Function ResolveActorPathForPartNo( _
    ByVal swDoc As DSComposer.Document, _
    ByVal targetPartNo As String, _
    ByVal uniqueAssemblies As Object, _
    ByRef actorsPart() As String) As String

    Dim actorPath As Variant
    Dim i As Long
    Dim candidatePath As String
    Dim selectionString As String
    Dim canonicalPartNo As String

    ResolveActorPathForPartNo = ""

    For Each actorPath In uniqueAssemblies.Keys
        candidatePath = CStr(actorPath)
        canonicalPartNo = GetCanonicalPartNo(swDoc, candidatePath)
        If StrComp(canonicalPartNo, targetPartNo, vbTextCompare) = 0 Then
            ResolveActorPathForPartNo = candidatePath
            Exit Function
        End If
    Next actorPath

    For i = 0 To UBound(actorsPart)
        candidatePath = Trim$(actorsPart(i))
        If Len(candidatePath) > 0 And IsPartActor(candidatePath) Then
            canonicalPartNo = GetCanonicalPartNo(swDoc, candidatePath)
            If StrComp(canonicalPartNo, targetPartNo, vbTextCompare) = 0 Then
                ResolveActorPathForPartNo = candidatePath
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub BuildActorCaches( _
    ByVal swDoc As DSComposer.Document, _
    ByVal uniqueAssemblies As Object, _
    ByRef actorsPart() As String, _
    ByVal targetPartMap As Object, _
    ByVal actorToPartMap As Object, _
    ByVal preferredActorMap As Object, _
    ByVal matchedPartMap As Object, _
    ByRef matchedActorCount As Long, _
    ByVal logFile As Object)

    Dim actorPath As Variant
    Dim i As Long

    matchedActorCount = 0

    For Each actorPath In uniqueAssemblies.Keys
        AddActorToCaches swDoc, CStr(actorPath), targetPartMap, actorToPartMap, preferredActorMap, matchedPartMap, matchedActorCount, logFile
    Next actorPath

    For i = 0 To UBound(actorsPart)
        If Len(Trim$(actorsPart(i))) > 0 And IsPartActor(actorsPart(i)) Then
            AddActorToCaches swDoc, Trim$(actorsPart(i)), targetPartMap, actorToPartMap, preferredActorMap, matchedPartMap, matchedActorCount, logFile
        End If
    Next i

    LogPrint logFile, Now & " - Actor cache complete: actorToPart=" & actorToPartMap.Count & _
                      ", preferredPartMatches=" & preferredActorMap.Count
End Sub

Private Sub AddActorToCaches( _
    ByVal swDoc As DSComposer.Document, _
    ByVal actorPath As String, _
    ByVal targetPartMap As Object, _
    ByVal actorToPartMap As Object, _
    ByVal preferredActorMap As Object, _
    ByVal matchedPartMap As Object, _
    ByRef matchedActorCount As Long, _
    ByVal logFile As Object)

    Dim canonicalPartNo As String

    canonicalPartNo = GetCanonicalPartNo(swDoc, actorPath)
    If Len(canonicalPartNo) = 0 Then Exit Sub

    If Not actorToPartMap.Exists(actorPath) Then
        actorToPartMap.Add actorPath, canonicalPartNo
    End If

    If targetPartMap.Exists(canonicalPartNo) Then
        matchedActorCount = matchedActorCount + 1

        If Not matchedPartMap.Exists(canonicalPartNo) Then
            matchedPartMap.Add canonicalPartNo, True
        End If

        ' Assemblies are processed first, so first hit is preferred.
        If Not preferredActorMap.Exists(canonicalPartNo) Then
            preferredActorMap.Add canonicalPartNo, actorPath
            LogVerbose logFile, Now & " - Preferred actor cached for part '" & canonicalPartNo & "': '" & actorPath & "'"
        End If
    End If
End Sub

Private Function PromptForPartNumberFile(ByVal logFile As Object) As String
    Dim picker As Object
    Dim selectedPath As String

    On Error Resume Next
    Set picker = Application.FileDialog(3) ' msoFileDialogFilePicker
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - File picker unavailable in this environment: " & Err.description
        Err.Clear
        selectedPath = InputBox("Enter full path to the line-separated part-number .txt file:", "Skycode - Part Number File")
        PromptForPartNumberFile = Trim$(selectedPath)
        Exit Function
    End If
    On Error GoTo 0

    With picker
        .Title = "Select Part Number Text File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        If .Show = -1 Then
            PromptForPartNumberFile = .SelectedItems(1)
        Else
            PromptForPartNumberFile = ""
        End If
    End With
End Function

Private Function GetCanonicalPartNo(ByVal swDoc As DSComposer.Document, ByVal actorPath As String) As String
    Dim valuePartNo As String
    Dim valueCfgUpper As String
    Dim valueCfgCamel As String
    Dim selectionString As String

    selectionString = Trim$(actorPath)
    If Len(selectionString) > 0 Then
        valuePartNo = StripMetaValue(swDoc.GetPropertyValue(selectionString, "Meta.PartNo"))
        valueCfgUpper = StripMetaValue(swDoc.GetPropertyValue(selectionString, "Meta.PARTNO_cfg"))
        valueCfgCamel = StripMetaValue(swDoc.GetPropertyValue(selectionString, "Meta.PartNo_cfg"))
    End If

    If Len(valuePartNo) = 0 And Len(valueCfgUpper) = 0 And Len(valueCfgCamel) = 0 Then
        If Left$(selectionString, 1) <> "<" Then
            selectionString = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"
            valuePartNo = StripMetaValue(swDoc.GetPropertyValue(selectionString, "Meta.PartNo"))
            valueCfgUpper = StripMetaValue(swDoc.GetPropertyValue(selectionString, "Meta.PARTNO_cfg"))
            valueCfgCamel = StripMetaValue(swDoc.GetPropertyValue(selectionString, "Meta.PartNo_cfg"))
        End If
    End If

    If Len(valueCfgUpper) > 0 Then
        GetCanonicalPartNo = valueCfgUpper
    ElseIf Len(valueCfgCamel) > 0 Then
        GetCanonicalPartNo = valueCfgCamel
    Else
        GetCanonicalPartNo = valuePartNo
    End If
End Function

Private Function TrySelectActorPath(ByVal swDoc As DSComposer.Document, ByVal actorPath As String, ByRef selectionString As String) As Boolean
    selectionString = Trim$(actorPath)
    If Len(selectionString) = 0 Then Exit Function

    If swDoc.SelectOnFilter(selectionString, False) Then
        TrySelectActorPath = True
        Exit Function
    End If

    If Left$(selectionString, 1) <> "<" Then
        selectionString = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"
        If swDoc.SelectOnFilter(selectionString, False) Then
            TrySelectActorPath = True
        End If
    End If
End Function

Private Function CollectUniqueAssemblyPaths(ByRef actors() As String, ByVal logFile As Object) As Object
    Dim map As Object
    Dim i As Long
    Dim j As Long
    Dim actorPath As String
    Dim segments() As String
    Dim assemblyPath As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    For i = 0 To UBound(actors)
        actorPath = Trim$(actors(i))
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            If UBound(segments) >= 0 Then
                assemblyPath = segments(0)
                If Not map.Exists(assemblyPath) Then map.Add assemblyPath, True

                If UBound(segments) > 0 Then
                    For j = 1 To UBound(segments) - 1
                        assemblyPath = assemblyPath & "." & segments(j)
                        If Not map.Exists(assemblyPath) Then map.Add assemblyPath, True
                    Next j
                End If
            End If
        End If
    Next i

    LogPrint logFile, Now & " - Unique assembly paths collected: " & map.Count
    Set CollectUniqueAssemblyPaths = map
End Function

Private Function IsPartActor(ByVal actorPath As String) As Boolean
    Dim segments() As String
    Dim lastSegment As String

    If Len(Trim$(actorPath)) = 0 Then
        IsPartActor = False
        Exit Function
    End If

    segments = Split(actorPath, ".")
    If UBound(segments) < 0 Then
        IsPartActor = False
        Exit Function
    End If

    lastSegment = segments(UBound(segments))
    IsPartActor = (Left$(lastSegment, 1) = "P")
End Function

Private Function StripMetaValue(ByVal rawValue As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim temp As String

    startPos = InStr(rawValue, "Value=""")
    If startPos > 0 Then
        startPos = startPos + Len("Value=""")
        endPos = InStr(startPos, rawValue, """")
        If endPos > startPos Then
            temp = Mid$(rawValue, startPos, endPos - startPos)
        End If
    End If

    If Len(temp) = 0 Then
        temp = rawValue
        temp = Replace(temp, "<Meta.PartNo>", "")
        temp = Replace(temp, "</Meta.PartNo>", "")
        temp = Replace(temp, "<Meta.PARTNO_cfg>", "")
        temp = Replace(temp, "</Meta.PARTNO_cfg>", "")
        temp = Replace(temp, "<Meta.PartNo_cfg>", "")
        temp = Replace(temp, "</Meta.PartNo_cfg>", "")
    End If

    temp = Replace(temp, Chr$(13), "")
    temp = Replace(temp, Chr$(10), "")
    temp = Replace(temp, Chr$(9), "")
    temp = Replace(temp, "&#13;", "")
    temp = Replace(temp, "&#10;", "")
    temp = Replace(temp, "&#9;", "")
    temp = Replace(temp, "&amp;", "&")

    StripMetaValue = Trim$(temp)
End Function

Private Function StripMetaPartNo(partNo As String) As String
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

Private Function BuildTargetNameMapFromTextFile(ByVal filePath As String, ByVal logFile As Object) As Object
    Dim fs As Object
    Dim textFile As Object
    Dim lineValue As String
    Dim map As Object

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set textFile = fs.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to open part-number file: " & Err.description
        Err.Clear
        Set BuildTargetNameMapFromTextFile = map
        Exit Function
    End If
    On Error GoTo 0

    Do While Not textFile.AtEndOfStream
        lineValue = Trim$(textFile.ReadLine)
        If Len(lineValue) > 0 Then
            If Not map.Exists(lineValue) Then
                map.Add lineValue, True
                LogVerbose logFile, Now & " - Target part number parsed: '" & lineValue & "'"
            End If
        End If
    Loop

    textFile.Close

    Set textFile = Nothing
    Set fs = Nothing
    Set BuildTargetNameMapFromTextFile = map
End Function

Private Function TryParseSelection(ByVal selectionText As String, ByRef names() As String, ByVal logFile As Object, ByVal label As String) As Boolean
    Dim payload As String
    Dim startPos As Long
    Dim endPos As Long

    payload = selectionText
    LogVerbose logFile, Now & " - Raw " & label & " selection (length " & Len(payload) & "): '" & payload & "'"

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
        TryParseSelection = False
        Exit Function
    End If

    names = Split(payload, "#")
    TryParseSelection = (UBound(names) >= 0)
End Function

Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub

Sub LogVerbose(logFile As Object, message As String)
    If VERBOSE_LOG Then
        LogPrint logFile, message
    End If
End Sub
