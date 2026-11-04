Option Compare Database
Option Explicit

' ============================================================
' Skycode - Set Meta PartNo from PartNo_cfg
' For every part (CLitActorMesh) and view (CLitActorView):
'   - Read Meta.PartNo
'   - Read Meta.PartNo_cfg
'   - If Meta.PartNo is blank or invalid -> update from Meta.PartNo_cfg
' ============================================================

Sub setMetaPartNoFromPartno_cfg()
    ' Declare variables
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim logFile As Object
    Dim fs As Object
    Dim actorPath As String
    Dim partPartNo As String
    Dim cleanPartPartNo As String
    Dim partNoCfg As String
    Dim cleanPartNoCfg As String
    Dim partViewName As String
    Dim i As Long
    Dim j As Long
    Dim char As String
    Dim actors() As String
    Dim actorSelection As String
    Dim processedParts As New Collection ' Track unique part numbers
    Dim segments() As String
    Dim partCount As Long ' Track processed parts
    Dim partNoCfgPropertyName As String ' Track which property name exists
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\meta partnos.txt", True)
    LogPrint logFile, Now & " - Starting part property fix test..."

    ' Connect to Composer
    On Error Resume Next
    Set swApp = New DSComposer.Application
    If Err.Number <> 0 Then
        LogPrint logFile, Now & " - Failed to create DSComposer app: " & Err.description
        Exit Sub
    End If
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - No active document!"
        Exit Sub
    End If
    On Error GoTo 0
    LogPrint logFile, Now & " - Connected to document"

    ' Get all mesh and view actors (parts)
    Dim actorSelectionView As String
    swDoc.AssySelectionMode = 0
    actorSelection = swDoc.GetAllActorsByType("CLitActorMesh")
    actorSelectionView = swDoc.GetAllActorsByType("CLitActorView")
    LogPrint logFile, Now & " - Mesh actors (length " & Len(actorSelection) & ") | View actors (length " & Len(actorSelectionView) & ")"
    
    ' Merge both actor lists
    Dim combinedSelection As String
    If Len(actorSelection) > 0 And Len(actorSelectionView) > 0 Then
        combinedSelection = actorSelection & "#" & actorSelectionView
    ElseIf Len(actorSelection) > 0 Then
        combinedSelection = actorSelection
    ElseIf Len(actorSelectionView) > 0 Then
        combinedSelection = actorSelectionView
    Else
        LogPrint logFile, Now & " - No actors found!"
        Exit Sub
    End If
    
    actors = Split(Replace(Replace(combinedSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    LogPrint logFile, Now & " - Found " & UBound(actors) + 1 & " total actors"
    
    ' Process parts with non-numeric Meta.PartNo
    partCount = 0
    For i = 0 To UBound(actors)
        On Error Resume Next
        actorPath = Trim(actors(i))
        If Len(actorPath) = 0 Then GoTo NextActor
        
        ' Clean up any leading CLitSelection tags
        If InStr(actorPath, ">") > 0 Then
            actorPath = Mid(actorPath, InStr(actorPath, ">") + 1)
        End If
        actorPath = Trim(actorPath)
        
        If Len(actorPath) > 0 Then
            ' Get last segment for use in SetPropertyMap Method 2
            Dim lastSegment As String
            segments = Split(actorPath, ".")
            lastSegment = segments(UBound(segments))
            
            ' Check Meta.PartNo
            Dim selectionString As String
            selectionString = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"
            partPartNo = swDoc.GetPropertyValue(selectionString, "Meta.PartNo")
            cleanPartPartNo = StripMetaPartNo(partPartNo)
            
            ' Get _cfg value first
            partNoCfg = swDoc.GetPropertyValue(selectionString, "Meta.PARTNO_cfg")
            cleanPartNoCfg = StripMetaPartNo(partNoCfg)
            
            If Len(cleanPartNoCfg) = 0 Then
                partNoCfg = swDoc.GetPropertyValue(selectionString, "Meta.PartNo_cfg")
                cleanPartNoCfg = StripMetaPartNo(partNoCfg)
                If Len(cleanPartNoCfg) > 0 Then
                    partNoCfgPropertyName = "Meta.PartNo_cfg"
                Else
                    partNoCfgPropertyName = ""
                End If
            Else
                partNoCfgPropertyName = "Meta.PARTNO_cfg"
            End If
            
            ' Process if _cfg has a value and it differs from current Meta.PartNo
            If Len(cleanPartNoCfg) > 0 And cleanPartNoCfg <> cleanPartPartNo Then
                partCount = partCount + 1
                LogPrint logFile, Now & " - Part #" & partCount & ": '" & actorPath & "' | Meta.PartNo='" & cleanPartPartNo & "' | " & partNoCfgPropertyName & "='" & cleanPartNoCfg & "'"
                
                ' Set Meta.PartNo from _cfg value
                If Len(cleanPartNoCfg) > 0 Then
                    Dim propertyMap As String
                    Dim propertySet As String
                    Dim success As Boolean
                    success = False
                    
                    ' Try Method 1: SetPropertyMap with full path name
                    propertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & actorPath & """><Meta.PartNo Value=""" & cleanPartNoCfg & """/></CLitModifiable></CLitPropertyMap>"
                    LogPrint logFile, Now & " - Method 1: Trying SetPropertyMap with full path: " & propertyMap
                    If swDoc.SetPropertyMap(propertyMap) Then
                        success = True
                        LogPrint logFile, Now & " - SUCCESS (Method 1): Set Meta.PartNo to '" & cleanPartNoCfg & "'"
                    Else
                        LogPrint logFile, Now & " - Method 1 failed, trying Method 2..."
                        
                        ' Try Method 2: SetPropertyMap with just last segment
                        propertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & lastSegment & """><Meta.PartNo Value=""" & cleanPartNoCfg & """/></CLitModifiable></CLitPropertyMap>"
                        LogPrint logFile, Now & " - Method 2: Trying SetPropertyMap with last segment: " & propertyMap
                        If swDoc.SetPropertyMap(propertyMap) Then
                            success = True
                            LogPrint logFile, Now & " - SUCCESS (Method 2): Set Meta.PartNo to '" & cleanPartNoCfg & "'"
                        Else
                            LogPrint logFile, Now & " - Method 2 failed, trying Method 3..."
                            
                            ' Try Method 3: SetPropertySet with selection string
                            propertySet = "<CLitPropertySet><Meta.PartNo Value=""" & cleanPartNoCfg & """/></CLitPropertySet>"
                            LogPrint logFile, Now & " - Method 3: Trying SetPropertySet: " & propertySet
                            If swDoc.SetPropertySet(selectionString, propertySet) Then
                                success = True
                                LogPrint logFile, Now & " - SUCCESS (Method 3): Set Meta.PartNo to '" & cleanPartNoCfg & "'"
                            Else
                                LogPrint logFile, Now & " - Method 3 failed. All methods exhausted for " & IIf(Len(partNoCfgPropertyName) > 0, partNoCfgPropertyName, "Meta.PARTNO_cfg") & "."
                            End If
                        End If
                    End If
                    
                    ' If all methods failed with Meta.PARTNO_cfg, try Meta.SWConfig
                    If Not success Then
                        LogPrint logFile, Now & " - FAILED: Could not set Meta.PartNo to '" & cleanPartNoCfg & "' for actor '" & actorPath & "' using any method"
                    End If
                End If
            ElseIf Len(cleanPartNoCfg) = 0 Then
                partCount = partCount + 1
                LogPrint logFile, Now & " - Part #" & partCount & ": '" & actorPath & "' | No _cfg value found, resetting Meta.PartNo to sanitized value '" & cleanPartPartNo & "'"

                Dim resetPropertyMap As String
                Dim resetPropertySet As String
                Dim resetSuccess As Boolean
                resetSuccess = False

                ' Try Method 1: SetPropertyMap with full path name
                resetPropertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & actorPath & """><Meta.PartNo Value=""" & cleanPartPartNo & """/></CLitModifiable></CLitPropertyMap>"
                LogPrint logFile, Now & " - Method 1: Trying SetPropertyMap with full path: " & resetPropertyMap
                If swDoc.SetPropertyMap(resetPropertyMap) Then
                    resetSuccess = True
                    LogPrint logFile, Now & " - SUCCESS (Method 1): Reset Meta.PartNo to '" & cleanPartPartNo & "'"
                Else
                    LogPrint logFile, Now & " - Method 1 failed, trying Method 2..."

                    ' Try Method 2: SetPropertyMap with just last segment
                    resetPropertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & lastSegment & """><Meta.PartNo Value=""" & cleanPartPartNo & """/></CLitModifiable></CLitPropertyMap>"
                    LogPrint logFile, Now & " - Method 2: Trying SetPropertyMap with last segment: " & resetPropertyMap
                    If swDoc.SetPropertyMap(resetPropertyMap) Then
                        resetSuccess = True
                        LogPrint logFile, Now & " - SUCCESS (Method 2): Reset Meta.PartNo to '" & cleanPartPartNo & "'"
                    Else
                        LogPrint logFile, Now & " - Method 2 failed, trying Method 3..."

                        ' Try Method 3: SetPropertySet with selection string
                        resetPropertySet = "<CLitPropertySet><Meta.PartNo Value=""" & cleanPartPartNo & """/></CLitPropertySet>"
                        LogPrint logFile, Now & " - Method 3: Trying SetPropertySet: " & resetPropertySet
                        If swDoc.SetPropertySet(selectionString, resetPropertySet) Then
                            resetSuccess = True
                            LogPrint logFile, Now & " - SUCCESS (Method 3): Reset Meta.PartNo to '" & cleanPartPartNo & "'"
                        Else
                            LogPrint logFile, Now & " - Method 3 failed. All methods exhausted while resetting Meta.PartNo."
                        End If
                    End If
                End If

                If Not resetSuccess Then
                    LogPrint logFile, Now & " - FAILED: Could not reset Meta.PartNo for actor '" & actorPath & "'"
                End If
            Else
                LogPrint logFile, Now & " - OK: Actor '" & actorPath & "' Meta.PartNo already matches _cfg ('" & cleanPartPartNo & "')"
            End If
        End If
NextActor:
        On Error GoTo 0
    Next i
    
    ' Log final part count
    LogPrint logFile, Now & " - Total parts processed: " & partCount
    
    ' Cleanup
    LogPrint logFile, Now & " - Test complete"
    logFile.Close
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

' Helper function to strip XML tags and extract the value
Function StripMetaPartNo(partNo As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(partNo, "Value=""")
    If startPos > 0 Then
        startPos = startPos + Len("Value=""")
        endPos = InStr(startPos, partNo, """")
        If endPos > startPos Then
            StripMetaPartNo = SanitizeMetaPartNo(Mid(partNo, startPos, endPos - startPos))
            Exit Function
        End If
    End If
    Dim temp As String
    temp = Replace(partNo, "<Meta.PartNo>", "")
    temp = Replace(temp, "</Meta.PartNo>", "")
    temp = SanitizeMetaPartNo(temp)
    
    ' Check if we got back the property name itself (means no value)
    If temp = "Meta.PartNo" Or temp = "Meta.PARTNO_cfg" Or temp = "Default" Then
        StripMetaPartNo = ""
    ElseIf Len(temp) > 0 Then
        StripMetaPartNo = temp
    Else
        StripMetaPartNo = ""
    End If
End Function

Function SanitizeMetaPartNo(value As String) As String
    Dim i As Long
    Dim ch As String
    Dim cleaned As String
    Dim re As Object

    ' Remove encoded XML newline/tab entities first (e.g. "&#13;") so
    ' we do not keep their digits as a false suffix like "13".
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "&#\d+;"
    cleaned = re.Replace(value, " ")

    cleaned = Replace(cleaned, "&quot;", """")
    cleaned = Replace(cleaned, "&apos;", "'")
    cleaned = Replace(cleaned, "&amp;", "&")

    cleaned = Replace(cleaned, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = Trim(cleaned)

    ' Keep only letters, numbers, space, dot, dash, and underscore.
    For i = 1 To Len(cleaned)
        ch = Mid(cleaned, i, 1)
        If ch Like "[A-Za-z0-9._ -]" Then
            SanitizeMetaPartNo = SanitizeMetaPartNo & ch
        End If
    Next i

    Do While InStr(SanitizeMetaPartNo, "  ") > 0
        SanitizeMetaPartNo = Replace(SanitizeMetaPartNo, "  ", " ")
    Loop
End Function

' Helper logging function
Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub

