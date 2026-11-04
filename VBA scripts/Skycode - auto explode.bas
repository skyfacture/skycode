Option Compare Database
Option Explicit

' ============================================================
' Skycode - Auto Explode (Active View, First-Level Candidates)
' Current behavior:
'   - Reads visible actors from the active view only.
'   - Detects root from visible paths and derives first-level assembly candidates.
'   - For each first-level candidate, builds a descendant parts selection.
'   - Applies LinearExplode on descendant parts using fixed +Z direction and distance.
'   - Writes per-phase diagnostics and summary to logs/auto explode.txt.
' Notes:
'   - This is a deterministic baseline implementation, not yet distance-search based.
' ============================================================

Private Const LOG_FILE_PATH As String = "C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\auto explode.txt"
Private Const MODE_PART_ACTORS As Integer = 0
Private Const MODE_VISIBLE_ACTORS As Integer = 1

Private Const EXPLODE_VECTOR_X As Integer = 0
Private Const EXPLODE_VECTOR_Y As Integer = 0
Private Const EXPLODE_VECTOR_Z As Integer = 1
Private Const EXPLODE_DISTANCE As Single = 400!

Public Sub ExplodeFirstLevelSubAssembliesActiveView()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim fs As Object
    Dim logFile As Object

    Dim actors() As String
    Dim rootID As String
    Dim firstLevelAssemblies As Collection
    Dim subAssemblyID As Variant

    Dim scannedCount As Long
    Dim candidateCount As Long
    Dim explodedCount As Long
    Dim skippedEmptyCount As Long
    Dim skippedNotFirstLevelCount As Long
    Dim skippedNotAssemblyPatternCount As Long
    Dim skippedNoPartMembersCount As Long
    Dim failCount As Long

    LogOpen fs, logFile, LOG_FILE_PATH
    LogPrint logFile, Now & " - ExplodeFirstLevelSubAssembliesActiveView starting..."

    ' Connect to Composer using the same host/runtime pattern used across Skycode macros.
    On Error Resume Next
    Set swApp = New DSComposer.Application
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - ERROR: Failed to create DSComposer.Application: " & Err.Description
        GoTo Cleanup
    End If

    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - ERROR: No active document."
        GoTo Cleanup
    End If
    On Error GoTo 0

    LogPrint logFile, Now & " - Phase 1: Collecting visible actors from active view..."
    If Not GetVisibleActorsFromActiveView(swDoc, actors, logFile) Then
        GoTo Summary
    End If
    rootID = DetectRootID(actors)

    If Len(rootID) = 0 Then
        LogPrint logFile, Now & " - ERROR: Could not detect root actor path from visible set."
        GoTo Summary
    End If

    LogPrint logFile, Now & " - Root detected: " & rootID
    LogPrint logFile, Now & " - Phase 2: Deriving first-level sub-assemblies from visible actor paths..."

    Set firstLevelAssemblies = DeriveFirstLevelAssemblyCandidates( _
        actors, _
        scannedCount, _
        skippedEmptyCount, _
        skippedNotFirstLevelCount, _
        skippedNotAssemblyPatternCount _
    )

    If firstLevelAssemblies.Count = 0 Then
        LogPrint logFile, Now & " - No first-level sub-assemblies derived from visible actors."
        GoTo Summary
    End If

    LogPrint logFile, Now & " - Derived first-level sub-assemblies: " & firstLevelAssemblies.Count
    LogPrint logFile, Now & " - Phase 3: Applying LinearExplode to descendant parts for each first-level sub-assembly..."

    ' Use part-level explode selection to simulate sub-assembly explode even when assembly explode is blocked.
    swDoc.AssySelectionMode = MODE_PART_ACTORS

    For Each subAssemblyID In firstLevelAssemblies
        actorPath = CStr(subAssemblyID)
        candidateCount = candidateCount + 1

        If TryLinearExplodeFromSubAssemblyParts(swDoc, actors, actorPath, logFile) Then
            explodedCount = explodedCount + 1
            LogPrint logFile, Now & " - Exploded: " & actorPath & " (vector Z=+1, distance=400)"
        Else
            failCount = failCount + 1
            If CountPartMembers(actors, actorPath) = 0 Then
                skippedNoPartMembersCount = skippedNoPartMembersCount + 1
                LogPrint logFile, Now & " - SKIPPED (no descendant parts): " & actorPath
            Else
                LogPrint logFile, Now & " - FAILED LinearExplode: " & actorPath
            End If
        End If
    Next subAssemblyID

Summary:
    LogPrint logFile, Now & " - Phase complete."
    LogPrint logFile, Now & " - Summary | Scanned=" & scannedCount & _
        " | Candidates=" & candidateCount & _
        " | Exploded=" & explodedCount & _
        " | Failed=" & failCount & _
        " | SkippedNoPartMembers=" & skippedNoPartMembersCount & _
        " | SkippedNotAssemblyPattern=" & skippedNotAssemblyPatternCount & _
        " | SkippedNotFirstLevel=" & skippedNotFirstLevelCount & _
        " | SkippedEmpty=" & skippedEmptyCount
    LogPrint logFile, Now & " - ExplodeFirstLevelSubAssembliesActiveView finished."

Cleanup:
    On Error Resume Next
    If Not logFile Is Nothing Then logFile.Close
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
    On Error GoTo 0
End Sub

Private Function GetVisibleActorsFromActiveView(ByRef swDoc As DSComposer.Document, ByRef actors() As String, ByRef logFile As Object) As Boolean
    Dim actorSelection As String
    Dim normalizedSelection As String

    swDoc.AssySelectionMode = MODE_VISIBLE_ACTORS
    actorSelection = swDoc.GetAllActors(MODE_VISIBLE_ACTORS) ' crAtvVisible = 1

    If Len(Trim(actorSelection)) = 0 Then
        LogPrint logFile, Now & " - No visible actors returned by GetAllActors(1)."
        GetVisibleActorsFromActiveView = False
        Exit Function
    End If

    normalizedSelection = NormalizeSelectionPayload(actorSelection)
    If Len(normalizedSelection) = 0 Then
        LogPrint logFile, Now & " - No parseable actor IDs in selection payload."
        GetVisibleActorsFromActiveView = False
        Exit Function
    End If

    actors = Split(normalizedSelection, "#")
    GetVisibleActorsFromActiveView = True
End Function

Private Function DeriveFirstLevelAssemblyCandidates(ByRef actors() As String, ByRef scannedCount As Long, ByRef skippedEmptyCount As Long, ByRef skippedNotFirstLevelCount As Long, ByRef skippedNotAssemblyPatternCount As Long) As Collection
    Dim i As Long
    Dim actorPath As String
    Dim segments() As String
    Dim firstLevelPath As String
    Dim result As Collection

    Set result = New Collection

    On Error Resume Next
    For i = 0 To UBound(actors)
        actorPath = CleanActorPath(actors(i))

        If Len(actorPath) = 0 Then
            skippedEmptyCount = skippedEmptyCount + 1
            GoTo NextVisibleActor
        End If

        scannedCount = scannedCount + 1
        segments = Split(actorPath, ".")

        If UBound(segments) < 1 Then
            skippedNotFirstLevelCount = skippedNotFirstLevelCount + 1
            GoTo NextVisibleActor
        End If

        ' First-level path is Root.Child, derived from any visible descendant under that child.
        firstLevelPath = Trim(segments(0)) & "." & Trim(segments(1))

        ' Existing Skycode convention: assembly nodes typically start with A, parts with P.
        If Left$(Trim(segments(1)), 1) <> "A" Then
            skippedNotAssemblyPatternCount = skippedNotAssemblyPatternCount + 1
            GoTo NextVisibleActor
        End If

        result.Add firstLevelPath, firstLevelPath
        Err.Clear

NextVisibleActor:
    Next i
    On Error GoTo 0

    Set DeriveFirstLevelAssemblyCandidates = result
End Function

Private Function TryLinearExplodeFromSubAssemblyParts(ByRef swDoc As DSComposer.Document, ByRef actors() As String, ByVal subAssemblyPath As String, ByRef logFile As Object) As Boolean
    Dim selectionByParts As String

    selectionByParts = BuildPartSelectionForSubAssembly(actors, subAssemblyPath)
    If Len(selectionByParts) = 0 Then
        TryLinearExplodeFromSubAssemblyParts = False
        Exit Function
    End If

    On Error Resume Next
    If swDoc.LinearExplode(selectionByParts, EXPLODE_VECTOR_X, EXPLODE_VECTOR_Y, EXPLODE_VECTOR_Z, EXPLODE_DISTANCE) Then
        TryLinearExplodeFromSubAssemblyParts = True
    Else
        TryLinearExplodeFromSubAssemblyParts = False
        If Err.Number <> 0 Then
            LogPrint logFile, Now & " - LinearExplode(parts selection) error on " & subAssemblyPath & ": " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

Private Function BuildPartSelectionForSubAssembly(ByRef actors() As String, ByVal subAssemblyPath As String) As String
    Dim i As Long
    Dim actorPath As String
    Dim selection As String

    selection = "<CLitSelection Name="""">"

    For i = 0 To UBound(actors)
        actorPath = CleanActorPath(actors(i))
        If IsDescendantPath(actorPath, subAssemblyPath) And IsPartPath(actorPath) Then
            selection = selection & "<CLitModifiable Name=""" & actorPath & """/>"
        End If
    Next i

    selection = selection & "</CLitSelection>"

    If selection = "<CLitSelection Name=""""></CLitSelection>" Then
        BuildPartSelectionForSubAssembly = ""
    Else
        BuildPartSelectionForSubAssembly = selection
    End If
End Function

Private Function CountPartMembers(ByRef actors() As String, ByVal subAssemblyPath As String) As Long
    Dim i As Long
    Dim actorPath As String

    For i = 0 To UBound(actors)
        actorPath = CleanActorPath(actors(i))
        If IsDescendantPath(actorPath, subAssemblyPath) And IsPartPath(actorPath) Then
            CountPartMembers = CountPartMembers + 1
        End If
    Next i
End Function

Private Function IsDescendantPath(ByVal actorPath As String, ByVal parentPath As String) As Boolean
    If Len(actorPath) = 0 Or Len(parentPath) = 0 Then
        IsDescendantPath = False
        Exit Function
    End If

    IsDescendantPath = (Left$(actorPath, Len(parentPath) + 1) = parentPath & ".")
End Function

Private Function IsPartPath(ByVal actorPath As String) As Boolean
    Dim segments() As String
    Dim leaf As String

    If Len(actorPath) = 0 Then
        IsPartPath = False
        Exit Function
    End If

    segments = Split(actorPath, ".")
    leaf = Trim(segments(UBound(segments)))
    IsPartPath = (Left$(leaf, 1) = "P")
End Function

Private Sub LogOpen(ByRef fs As Object, ByRef logFile As Object, ByVal outputPath As String)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile(outputPath, True)
End Sub

Private Function NormalizeSelectionPayload(ByVal payload As String) As String
    Dim result As String

    result = payload
    result = Replace(result, "<CLitSelection Name="""" Ident=""-1"">", "")
    result = Replace(result, "<CLitSelection Name="""">", "")
    result = Replace(result, "</CLitSelection>", "")

    NormalizeSelectionPayload = Trim(result)
End Function

Private Function DetectRootID(ByRef actors() As String) As String
    Dim i As Long
    Dim actorPath As String
    Dim segments() As String

    For i = 0 To UBound(actors)
        actorPath = CleanActorPath(actors(i))
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            If UBound(segments) >= 0 Then
                DetectRootID = Trim(segments(0))
                Exit Function
            End If
        End If
    Next i

    DetectRootID = ""
End Function


Private Function CleanActorPath(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Trim(value)

    If InStr(cleaned, "<CLitModifiable") > 0 Then
        cleaned = Replace(cleaned, "<CLitModifiable Name=""", "")
        cleaned = Replace(cleaned, """/>", "")
    End If

    If InStr(cleaned, ">") > 0 Then
        cleaned = Mid(cleaned, InStr(cleaned, ">") + 1)
    End If

    cleaned = Trim(cleaned)
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    cleaned = Replace(cleaned, vbTab, "")

    CleanActorPath = cleaned
End Function

Private Sub LogPrint(ByVal logFile As Object, ByVal message As String)
    Debug.Print message
    If Not logFile Is Nothing Then
        logFile.WriteLine message
    End If
End Sub
