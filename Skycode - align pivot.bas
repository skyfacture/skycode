Option Compare Database
Option Explicit

' ============================================================
' Skycode - Align Pivot to Actor Location Frame
' For every part (CLitActorMesh) and sub-assembly (CLitActorAssembly):
'   - Read Actor.Location Frame.Origin  (X, Y, Z)
'   - Read Actor.Location Pivot.Origin  (X, Y, Z)
'   - If Pivot.Origin already equals Frame.Origin  -> skip
'   - Else set Pivot.Origin = Frame.Origin
' ============================================================

Sub AlignPivotToFrame()

    ' ---- Declarations -------------------------------------------------------
    Dim swApp       As DSComposer.Application
    Dim swDoc       As DSComposer.Document
    Dim fs          As Object
    Dim logFile     As Object

    Dim actorSelection  As String
    Dim actors()        As String
    Dim actorPath       As String
    Dim selStr          As String
    Dim locationXml     As String

    Dim frameX As Double, frameY As Double, frameZ As Double
    Dim pivotX As Double, pivotY As Double, pivotZ As Double

    Dim frameVXX As String, frameVXY As String, frameVXZ As String
    Dim frameVYX As String, frameVYY As String, frameVYZ As String
    Dim pivotVXX As String, pivotVXY As String, pivotVXZ As String
    Dim pivotVYX As String, pivotVYY As String, pivotVYZ As String

    Dim i           As Long
    Dim skipCount   As Long
    Dim updateCount As Long
    Dim errorCount  As Long
    Const TOLERANCE As Double = 0.0001

    ' ---- Logging ------------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\align pivot.txt", True)
    LogMsg logFile, Now & " - AlignPivotToFrame starting..."

    ' ---- Connect to Composer ------------------------------------------------
    On Error Resume Next
    Set swApp = New DSComposer.Application
    If Err.Number <> 0 Then
        LogMsg logFile, "ERROR: Could not connect to DSComposer.Application - " & Err.Description
        GoTo Cleanup
    End If
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogMsg logFile, "ERROR: No active document."
        GoTo Cleanup
    End If
    On Error GoTo 0
    LogMsg logFile, "Connected to document."

    ' ---- Collect all parts and sub-assemblies -------------------------------
    ' AssySelectionMode 2 = select into sub-assemblies (gets all levels)
    swDoc.AssySelectionMode = 2

    Dim allActorPaths() As String
    Dim allCount As Long
    allCount = 0
    ReDim allActorPaths(0)

    ' --- Mesh actors (parts) ---
    actorSelection = swDoc.GetAllActorsByType("CLitActorMesh")
    If Len(actorSelection) > 0 Then
        Dim meshRaw As String
        meshRaw = Replace(actorSelection, "<CLitSelection Name="""" Ident=""-1"">", "")
        meshRaw = Replace(meshRaw, "</CLitSelection>", "")
        actors = Split(meshRaw, "#")
        Dim m As Long
        For m = 0 To UBound(actors)
            Dim mp As String
            mp = CleanActorPath(actors(m))
            If Len(mp) > 0 Then
                ReDim Preserve allActorPaths(allCount)
                allActorPaths(allCount) = mp
                allCount = allCount + 1
            End If
        Next m
        LogMsg logFile, "Mesh actors found: " & allCount
    Else
        LogMsg logFile, "No mesh actors found."
    End If

    ' --- Assembly actors (sub-assemblies) ---
    actorSelection = swDoc.GetAllActorsByType("CLitActorAssembly")
    Dim assyStart As Long
    assyStart = allCount
    If Len(actorSelection) > 0 Then
        Dim assyRaw As String
        assyRaw = Replace(actorSelection, "<CLitSelection Name="""" Ident=""-1"">", "")
        assyRaw = Replace(assyRaw, "</CLitSelection>", "")
        actors = Split(assyRaw, "#")
        Dim a As Long
        For a = 0 To UBound(actors)
            Dim ap As String
            ap = CleanActorPath(actors(a))
            If Len(ap) > 0 Then
                ReDim Preserve allActorPaths(allCount)
                allActorPaths(allCount) = ap
                allCount = allCount + 1
            End If
        Next a
        LogMsg logFile, "Assembly actors found: " & (allCount - assyStart)
    Else
        LogMsg logFile, "No assembly actors found."
    End If

    LogMsg logFile, "Total actors to process: " & allCount

    If allCount = 0 Then
        LogMsg logFile, "Nothing to process. Exiting."
        GoTo Cleanup
    End If

    ' ---- Process each actor -------------------------------------------------
    skipCount = 0
    updateCount = 0
    errorCount = 0

    For i = 0 To allCount - 1
        actorPath = allActorPaths(i)
        selStr = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"

        On Error Resume Next

        ' Get Actor.Location XML
        locationXml = swDoc.GetPropertyValue(selStr, "Actor.Location")

        If Err.Number <> 0 Or Len(Trim(locationXml)) = 0 Then
            LogMsg logFile, "  SKIP (no location): " & actorPath
            errorCount = errorCount + 1
            Err.Clear
            On Error GoTo 0
            GoTo NextActor
        End If
        On Error GoTo 0

        ' --- Parse Frame.Origin ----------------------------------------------
        Dim frameOriginTag As String
        frameOriginTag = ExtractTagContent(locationXml, "Frame")
        If Len(frameOriginTag) = 0 Then
            LogMsg logFile, "  SKIP (no Frame): " & actorPath
            errorCount = errorCount + 1
            GoTo NextActor
        End If

        ' Frame Origin
        Dim frameOriginStr As String
        frameOriginStr = ExtractFirstTag(frameOriginTag, "Origin")
        If Len(frameOriginStr) = 0 Then
            LogMsg logFile, "  SKIP (no Frame/Origin): " & actorPath
            errorCount = errorCount + 1
            GoTo NextActor
        End If
        frameX = CDbl(GetAttrValue(frameOriginStr, "X"))
        frameY = CDbl(GetAttrValue(frameOriginStr, "Y"))
        frameZ = CDbl(GetAttrValue(frameOriginStr, "Z"))

        ' Frame VX / VY (preserve when rebuilding)
        Dim frameVXStr As String, frameVYStr As String
        frameVXStr = ExtractFirstTag(frameOriginTag, "VX")
        frameVYStr = ExtractFirstTag(frameOriginTag, "VY")
        frameVXX = GetAttrValue(frameVXStr, "X") : If frameVXX = "" Then frameVXX = "1"
        frameVXY = GetAttrValue(frameVXStr, "Y") : If frameVXY = "" Then frameVXY = "0"
        frameVXZ = GetAttrValue(frameVXStr, "Z") : If frameVXZ = "" Then frameVXZ = "0"
        frameVYX = GetAttrValue(frameVYStr, "X") : If frameVYX = "" Then frameVYX = "0"
        frameVYY = GetAttrValue(frameVYStr, "Y") : If frameVYY = "" Then frameVYY = "1"
        frameVYZ = GetAttrValue(frameVYStr, "Z") : If frameVYZ = "" Then frameVYZ = "0"

        ' --- Parse Pivot.Origin ----------------------------------------------
        Dim pivotTag As String
        pivotTag = ExtractTagContent(locationXml, "Pivot")
        If Len(pivotTag) = 0 Then
            LogMsg logFile, "  SKIP (no Pivot): " & actorPath
            errorCount = errorCount + 1
            GoTo NextActor
        End If

        Dim pivotOriginStr As String
        pivotOriginStr = ExtractFirstTag(pivotTag, "Origin")
        If Len(pivotOriginStr) = 0 Then
            LogMsg logFile, "  SKIP (no Pivot/Origin): " & actorPath
            errorCount = errorCount + 1
            GoTo NextActor
        End If
        pivotX = CDbl(GetAttrValue(pivotOriginStr, "X"))
        pivotY = CDbl(GetAttrValue(pivotOriginStr, "Y"))
        pivotZ = CDbl(GetAttrValue(pivotOriginStr, "Z"))

        ' Pivot VX / VY (preserve when rebuilding)
        Dim pivotVXStr As String, pivotVYStr As String
        pivotVXStr = ExtractFirstTag(pivotTag, "VX")
        pivotVYStr = ExtractFirstTag(pivotTag, "VY")
        pivotVXX = GetAttrValue(pivotVXStr, "X") : If pivotVXX = "" Then pivotVXX = "1"
        pivotVXY = GetAttrValue(pivotVXStr, "Y") : If pivotVXY = "" Then pivotVXY = "0"
        pivotVXZ = GetAttrValue(pivotVXStr, "Z") : If pivotVXZ = "" Then pivotVXZ = "0"
        pivotVYX = GetAttrValue(pivotVYStr, "X") : If pivotVYX = "" Then pivotVYX = "0"
        pivotVYY = GetAttrValue(pivotVYStr, "Y") : If pivotVYY = "" Then pivotVYY = "1"
        pivotVYZ = GetAttrValue(pivotVYStr, "Z") : If pivotVYZ = "" Then pivotVYZ = "0"

        ' --- Compare (within tolerance) --------------------------------------
        Dim same As Boolean
        same = (Abs(pivotX - frameX) < TOLERANCE) And _
               (Abs(pivotY - frameY) < TOLERANCE) And _
               (Abs(pivotZ - frameZ) < TOLERANCE)

        If same Then
            LogMsg logFile, "  SKIP (already aligned) [" & FormatCoord(frameX) & ", " & FormatCoord(frameY) & ", " & FormatCoord(frameZ) & "]: " & actorPath
            skipCount = skipCount + 1
            GoTo NextActor
        End If

        LogMsg logFile, "  UPDATE: " & actorPath
        LogMsg logFile, "    Frame  Origin: " & FormatCoord(frameX) & ", " & FormatCoord(frameY) & ", " & FormatCoord(frameZ)
        LogMsg logFile, "    Pivot  Origin: " & FormatCoord(pivotX) & ", " & FormatCoord(pivotY) & ", " & FormatCoord(pivotZ)

        ' --- Build new Actor.Location XML with Pivot = Frame origin ----------
        Dim newLocation As String
        newLocation = "<Actor.Location>" & _
            "<Frame>" & _
                "<Origin X=""" & FormatCoord(frameX) & """ Y=""" & FormatCoord(frameY) & """ Z=""" & FormatCoord(frameZ) & """/>" & _
                "<VX X=""" & frameVXX & """ Y=""" & frameVXY & """ Z=""" & frameVXZ & """/>" & _
                "<VY X=""" & frameVYX & """ Y=""" & frameVYY & """ Z=""" & frameVYZ & """/>" & _
            "</Frame>" & _
            "<Pivot>" & _
                "<Origin X=""" & FormatCoord(frameX) & """ Y=""" & FormatCoord(frameY) & """ Z=""" & FormatCoord(frameZ) & """/>" & _
                "<VX X=""" & pivotVXX & """ Y=""" & pivotVXY & """ Z=""" & pivotVXZ & """/>" & _
                "<VY X=""" & pivotVYX & """ Y=""" & pivotVYY & """ Z=""" & pivotVYZ & """/>" & _
            "</Pivot>" & _
        "</Actor.Location>"

        ' --- Apply via SetPropertyMap ----------------------------------------
        Dim propMap As String
        propMap = "<CLitPropertyMap><CLitModifiable Name=""" & actorPath & """>" & newLocation & "</CLitModifiable></CLitPropertyMap>"

        On Error Resume Next
        Dim ok As Boolean
        ok = swDoc.SetPropertyMap(propMap)
        If Err.Number <> 0 Then
            LogMsg logFile, "    ERROR (SetPropertyMap exception): " & Err.Description
            Err.Clear
            On Error GoTo 0
            errorCount = errorCount + 1
            GoTo NextActor
        End If
        On Error GoTo 0

        If ok Then
            LogMsg logFile, "    -> SUCCESS: Pivot set to [" & FormatCoord(frameX) & ", " & FormatCoord(frameY) & ", " & FormatCoord(frameZ) & "]"
            updateCount = updateCount + 1
        Else
            LogMsg logFile, "    -> FAILED: SetPropertyMap returned False"
            errorCount = errorCount + 1
        End If

NextActor:
    Next i

    ' ---- Summary ------------------------------------------------------------
    LogMsg logFile, ""
    LogMsg logFile, "=== DONE ==="
    LogMsg logFile, "  Updated : " & updateCount
    LogMsg logFile, "  Skipped : " & skipCount
    LogMsg logFile, "  Errors  : " & errorCount
    LogMsg logFile, "  Total   : " & allCount

Cleanup:
    If Not logFile Is Nothing Then
        logFile.Close
        Set logFile = Nothing
    End If
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

' =============================================================================
' Helper: Remove leading XML wrapper / whitespace from a raw actor path token
' =============================================================================
Function CleanActorPath(raw As String) As String
    Dim s As String
    s = Trim(raw)
    ' Strip any leftover <CLitSelection ...> tag prefix
    If InStr(s, ">") > 0 Then
        s = Trim(Mid(s, InStr(s, ">") + 1))
    End If
    CleanActorPath = s
End Function

' =============================================================================
' Helper: Extract the inner content of the FIRST occurrence of <tagName ...>
'         or <tagName>...</tagName> from xml.
'         For self-closing tags returns the full tag string.
' =============================================================================
Function ExtractTagContent(xml As String, tagName As String) As String
    Dim openTag As String
    Dim startPos As Long, endPos As Long, contentStart As Long

    ' Try <tagName> ... </tagName>  (block tag)
    openTag = "<" & tagName & ">"
    startPos = InStr(xml, openTag)
    If startPos > 0 Then
        contentStart = startPos + Len(openTag)
        endPos = InStr(contentStart, xml, "</" & tagName & ">")
        If endPos > 0 Then
            ExtractTagContent = Mid(xml, contentStart, endPos - contentStart)
            Exit Function
        End If
    End If

    ExtractTagContent = ""
End Function

' =============================================================================
' Helper: Extract the first self-closing or opening tag <tagName ... />
'         Returns the full tag string including attributes (without the </>)
' =============================================================================
Function ExtractFirstTag(xml As String, tagName As String) As String
    Dim openStr As String
    Dim pos As Long, endPos As Long

    openStr = "<" & tagName & " "
    pos = InStr(xml, openStr)
    If pos = 0 Then
        ' Try exact tag with no attributes
        openStr = "<" & tagName & "/"
        pos = InStr(xml, openStr)
        If pos = 0 Then
            ExtractFirstTag = ""
            Exit Function
        End If
    End If

    ' Find closing /> or >
    endPos = InStr(pos, xml, "/>")
    If endPos = 0 Then endPos = InStr(pos, xml, ">")
    If endPos = 0 Then
        ExtractFirstTag = ""
        Exit Function
    End If

    ExtractFirstTag = Mid(xml, pos, endPos - pos + 2)
End Function

' =============================================================================
' Helper: Get a named attribute value from a tag string
'         e.g. GetAttrValue("<Origin X=""1.5"" Y=""0"" Z=""-3""/>", "X") -> "1.5"
' =============================================================================
Function GetAttrValue(tagStr As String, attrName As String) As String
    Dim search As String
    Dim pos As Long, valStart As Long, valEnd As Long

    search = attrName & "="""
    pos = InStr(tagStr, search)
    If pos = 0 Then
        GetAttrValue = ""
        Exit Function
    End If
    valStart = pos + Len(search)
    valEnd = InStr(valStart, tagStr, """")
    If valEnd <= valStart Then
        GetAttrValue = ""
        Exit Function
    End If
    GetAttrValue = Mid(tagStr, valStart, valEnd - valStart)
End Function

' =============================================================================
' Helper: Format a coordinate to a clean string (no trailing .0 noise)
' =============================================================================
Function FormatCoord(v As Double) As String
    ' Round to 6 decimal places, then strip trailing zeros
    Dim s As String
    s = Format(v, "0.######")
    ' Ensure no exponent notation slips through for very small/large values
    FormatCoord = s
End Function

' =============================================================================
' Helper: Write to log file and immediate window
' =============================================================================
Sub LogMsg(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub
