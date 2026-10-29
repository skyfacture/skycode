Option Compare Database
Option Explicit

' ============================================================
' Skycode - Set Meta Description from Description_cfg
' For every part (CLitActorMesh) and view (CLitActorView):
'   - Read Meta.Description
'   - Read Meta.Description_cfg
'   - If Meta.Description is blank or invalid -> update from Meta.Description_cfg
' ============================================================

Sub SetMetaDescriptionfromDescription_cfg()
    ' Declare variables
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim logFile As Object
    Dim fs As Object
    Dim actorPath As String
    Dim partDescription As String
    Dim cleanPartDescription As String
    Dim descriptionCfg As String
    Dim cleanDescriptionCfg As String
    Dim i As Long
    Dim actors() As String
    Dim actorSelection As String
    Dim segments() As String
    Dim partCount As Long ' Track processed parts
    Dim isPart As Boolean
    Dim lastSegment As String
    Dim selectionString As String
    Dim descPropertyMap As String
    Dim descPropertySet As String
    Dim descSuccess As Boolean
    Dim descriptionCfgPropertyName As String ' Track which property name exists
    Dim shouldUpdateDescription As Boolean
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\meta descriptions.txt", True)
    LogPrint logFile, Now & " - Starting description fix test..."

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
    
    ' Process parts with blank descriptions
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
            segments = Split(actorPath, ".")
            lastSegment = segments(UBound(segments))
            
            ' Process Meta.Description if blank or invalid
            selectionString = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"
            partDescription = swDoc.GetPropertyValue(selectionString, "Meta.Description")
            cleanPartDescription = StripMetaProperty(partDescription, "Meta.Description")
            cleanPartDescription = SanitizeMetaDescription(cleanPartDescription)
            
            ' Get Meta.Description_cfg and Meta.DESCRIPTION_cfg (both queried upfront)
            descriptionCfg = swDoc.GetPropertyValue(selectionString, "Meta.Description_cfg")
            cleanDescriptionCfg = StripMetaProperty(descriptionCfg, "Meta.Description_cfg")
            
            Dim cleanDescriptionCfgUpper As String
            Dim descriptionCfgUpper As String
            descriptionCfgUpper = swDoc.GetPropertyValue(selectionString, "Meta.DESCRIPTION_cfg")
            cleanDescriptionCfgUpper = StripMetaProperty(descriptionCfgUpper, "Meta.DESCRIPTION_cfg")
            
            ' Use whichever has a value (prefer Meta.Description_cfg, fall back to Meta.DESCRIPTION_cfg)
            If Len(cleanDescriptionCfg) > 0 Then
                descriptionCfgPropertyName = "Meta.Description_cfg"
                LogPrint logFile, Now & " - Meta.Description_cfg value: '" & cleanDescriptionCfg & "'"
            ElseIf Len(cleanDescriptionCfgUpper) > 0 Then
                cleanDescriptionCfg = cleanDescriptionCfgUpper
                descriptionCfgPropertyName = "Meta.DESCRIPTION_cfg"
                LogPrint logFile, Now & " - Meta.DESCRIPTION_cfg value: '" & cleanDescriptionCfg & "'"
            Else
                descriptionCfgPropertyName = ""
                LogPrint logFile, Now & " - Meta.Description_cfg value: '" & cleanDescriptionCfg & "'"
                LogPrint logFile, Now & " - Meta.DESCRIPTION_cfg value: '" & cleanDescriptionCfgUpper & "'"
            End If
            
            ' Clean up XML-escaped apostrophes, then sanitize.
            cleanDescriptionCfg = Replace(cleanDescriptionCfg, "&apos;", "/")
            cleanDescriptionCfg = SanitizeMetaDescription(cleanDescriptionCfg)

            shouldUpdateDescription = False

            If Len(cleanDescriptionCfg) > 0 Then
                If cleanPartDescription <> cleanDescriptionCfg Then
                    shouldUpdateDescription = True
                    partCount = partCount + 1

                    If Len(cleanPartDescription) = 0 Then
                        LogPrint logFile, Now & " - BLANK Description for part #" & partCount & ": '" & actorPath & "'"
                    ElseIf Not IsValidDescription(cleanPartDescription) Then
                        LogPrint logFile, Now & " - INVALID Description for part #" & partCount & ": '" & actorPath & "' | Current: '" & cleanPartDescription & "'"
                    Else
                        LogPrint logFile, Now & " - MISMATCH Description for part #" & partCount & ": '" & actorPath & "' | Current: '" & cleanPartDescription & "' | Canonical: '" & cleanDescriptionCfg & "'"
                    End If
                End If
            ElseIf Len(cleanPartDescription) = 0 Or Not IsValidDescription(cleanPartDescription) Then
                shouldUpdateDescription = True
                partCount = partCount + 1

                If Len(cleanPartDescription) = 0 Then
                    LogPrint logFile, Now & " - BLANK Description for part #" & partCount & ": '" & actorPath & "'"
                Else
                    LogPrint logFile, Now & " - INVALID Description for part #" & partCount & ": '" & actorPath & "' | Current: '" & cleanPartDescription & "'"
                End If
            End If

            If shouldUpdateDescription Then
                ' Set Meta.Description if we have a canonical non-Meta.Description value
                If Len(cleanDescriptionCfg) > 0 Then
                    descSuccess = False
                    
                    ' Try Method 1: SetPropertyMap with full path name
                    descPropertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & actorPath & """><Meta.Description Value=""" & cleanDescriptionCfg & """/></CLitModifiable></CLitPropertyMap>"
                    LogPrint logFile, Now & " - Method 1: Trying SetPropertyMap for Description with full path: " & descPropertyMap
                    If swDoc.SetPropertyMap(descPropertyMap) Then
                        descSuccess = True
                        LogPrint logFile, Now & " - SUCCESS (Method 1): Set Meta.Description to '" & cleanDescriptionCfg & "'"
                    Else
                        LogPrint logFile, Now & " - Method 1 failed, trying Method 2..."
                        
                        ' Try Method 2: SetPropertyMap with just last segment
                        descPropertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & lastSegment & """><Meta.Description Value=""" & cleanDescriptionCfg & """/></CLitModifiable></CLitPropertyMap>"
                        LogPrint logFile, Now & " - Method 2: Trying SetPropertyMap for Description with last segment: " & descPropertyMap
                        If swDoc.SetPropertyMap(descPropertyMap) Then
                            descSuccess = True
                            LogPrint logFile, Now & " - SUCCESS (Method 2): Set Meta.Description to '" & cleanDescriptionCfg & "'"
                        Else
                            LogPrint logFile, Now & " - Method 2 failed, trying Method 3..."
                            
                            ' Try Method 3: SetPropertySet with selection string
                            descPropertySet = "<CLitPropertySet><Meta.Description Value=""" & cleanDescriptionCfg & """/></CLitPropertySet>"
                            LogPrint logFile, Now & " - Method 3: Trying SetPropertySet for Description: " & descPropertySet
                            If swDoc.SetPropertySet(selectionString, descPropertySet) Then
                                descSuccess = True
                                LogPrint logFile, Now & " - SUCCESS (Method 3): Set Meta.Description to '" & cleanDescriptionCfg & "'"
                            Else
                                LogPrint logFile, Now & " - Method 3 failed. All methods exhausted for Meta.Description_cfg."
                            End If
                        End If
                    End If
                    
                    If Not descSuccess Then
                        LogPrint logFile, Now & " - FAILED: Could not set Meta.Description to '" & cleanDescriptionCfg & "' for actor '" & actorPath & "' using any method"
                    End If
                Else
                    LogPrint logFile, Now & " - No " & IIf(Len(descriptionCfgPropertyName) > 0, descriptionCfgPropertyName, "Meta.Description_cfg") & " value; resetting Meta.Description to sanitized value '" & cleanPartDescription & "'"

                    descSuccess = False

                    ' Try Method 1: SetPropertyMap with full path name
                    descPropertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & actorPath & """><Meta.Description Value=""" & cleanPartDescription & """/></CLitModifiable></CLitPropertyMap>"
                    LogPrint logFile, Now & " - Method 1: Trying SetPropertyMap for Description reset with full path: " & descPropertyMap
                    If swDoc.SetPropertyMap(descPropertyMap) Then
                        descSuccess = True
                        LogPrint logFile, Now & " - SUCCESS (Method 1): Reset Meta.Description to '" & cleanPartDescription & "'"
                    Else
                        LogPrint logFile, Now & " - Method 1 failed, trying Method 2..."

                        ' Try Method 2: SetPropertyMap with just last segment
                        descPropertyMap = "<CLitPropertyMap><CLitModifiable Name=""" & lastSegment & """><Meta.Description Value=""" & cleanPartDescription & """/></CLitModifiable></CLitPropertyMap>"
                        LogPrint logFile, Now & " - Method 2: Trying SetPropertyMap for Description reset with last segment: " & descPropertyMap
                        If swDoc.SetPropertyMap(descPropertyMap) Then
                            descSuccess = True
                            LogPrint logFile, Now & " - SUCCESS (Method 2): Reset Meta.Description to '" & cleanPartDescription & "'"
                        Else
                            LogPrint logFile, Now & " - Method 2 failed, trying Method 3..."

                            ' Try Method 3: SetPropertySet with selection string
                            descPropertySet = "<CLitPropertySet><Meta.Description Value=""" & cleanPartDescription & """/></CLitPropertySet>"
                            LogPrint logFile, Now & " - Method 3: Trying SetPropertySet for Description reset: " & descPropertySet
                            If swDoc.SetPropertySet(selectionString, descPropertySet) Then
                                descSuccess = True
                                LogPrint logFile, Now & " - SUCCESS (Method 3): Reset Meta.Description to '" & cleanPartDescription & "'"
                            Else
                                LogPrint logFile, Now & " - Method 3 failed. All methods exhausted while resetting Meta.Description."
                            End If
                        End If
                    End If

                    If Not descSuccess Then
                        LogPrint logFile, Now & " - FAILED: Could not reset Meta.Description for actor '" & actorPath & "'"
                    End If
                End If
            Else
                LogPrint logFile, Now & " - Meta.Description already has value: '" & cleanPartDescription & "'"
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

' Function to check if a description is valid
' Valid descriptions contain numbers, brackets, or special technical characters
' Invalid descriptions only contain letters, spaces, and hyphens
Function IsValidDescription(description As String) As Boolean
    Dim i As Long
    Dim char As String
    
    ' Empty descriptions are invalid
    If Len(description) = 0 Then
        IsValidDescription = False
        Exit Function
    End If
    
    ' Check each character
    For i = 1 To Len(description)
        char = Mid(description, i, 1)
        
        ' Check for numbers
        If char >= "0" And char <= "9" Then
            IsValidDescription = True
            Exit Function
        End If
        
        ' Check for brackets and parentheses
        If char = "(" Or char = ")" Or char = "[" Or char = "]" Or char = "{" Or char = "}" Then
            IsValidDescription = True
            Exit Function
        End If
        
        ' Check for technical special characters (not just letters, spaces, hyphens)
        ' Includes: forward slash, backslash, comma, diameter symbol, degree, multiplication
        If char = "/" Or char = "\" Or char = "," Or Asc(char) > 127 Then
            IsValidDescription = True
            Exit Function
        End If
        
        ' Check for period or dot (but not trailing period only)
        If char = "." And i < Len(description) Then
            IsValidDescription = True
            Exit Function
        End If
    Next i
    
    ' If we got here, description only contains letters, spaces, and hyphens
    IsValidDescription = False
End Function

' Generic helper function to strip XML tags and extract the value for any Meta property
Function StripMetaProperty(propValue As String, propName As String) As String
    Dim startPos As Long, endPos As Long
    startPos = InStr(propValue, "Value=""")
    If startPos > 0 Then
        startPos = startPos + Len("Value=""")
        endPos = InStr(startPos, propValue, """")
        If endPos > startPos Then
            StripMetaProperty = Trim(Mid(propValue, startPos, endPos - startPos))
            ' Clean up multiple spaces
            Do While InStr(StripMetaProperty, "  ") > 0
                StripMetaProperty = Replace(StripMetaProperty, "  ", " ")
            Loop
            Exit Function
        End If
    End If
    Dim temp As String
    temp = Replace(propValue, "<" & propName & ">", "")
    temp = Replace(temp, "</" & propName & ">", "")
    temp = Trim(temp)
    
    ' Clean up multiple spaces before validation
    Do While InStr(temp, "  ") > 0
        temp = Replace(temp, "  ", " ")
    Loop
    
    ' Check if we got back the property name itself (means no value)
    If temp = propName Or temp = propName & "_cfg" Or temp = "Default" Then
        StripMetaProperty = ""
    ElseIf Len(temp) > 0 Then
        StripMetaProperty = temp
    Else
        StripMetaProperty = ""
    End If
End Function

Function SanitizeMetaDescription(value As String) As String
    Dim cleaned As String

    cleaned = Replace(value, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = Trim(cleaned)

    ' Preserve the source description text and only normalize whitespace.
    SanitizeMetaDescription = cleaned

    Do While InStr(SanitizeMetaDescription, "  ") > 0
        SanitizeMetaDescription = Replace(SanitizeMetaDescription, "  ", " ")
    Loop
End Function

' Helper logging function
Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    logFile.WriteLine message
End Sub

