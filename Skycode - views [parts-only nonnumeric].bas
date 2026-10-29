Option Compare Database
Option Explicit

' ============================================================
' Skycode - Create Views for Non-Numeric Part Numbers
' For actors in part selection mode:
'   - Read Meta.PartNo
'   - Treat blank or non-numeric Meta.PartNo as candidates
'   - Prefer Meta.PARTNO_cfg, then Meta.PartNo_cfg as the canonical value
'   - Use that value as the view name (with fallback when empty)
'   - Skip duplicate view names
' ============================================================

Sub CreateAssemblyViews()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actors() As String
    Dim i As Long, j As Long
    Dim logFile As Object, fs As Object
    Dim actorPath As String
    Dim segments() As String
    Dim char As String ' For numeric check
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\views [parts-only nonnumeric].txt", True)
    LogPrint logFile, Now & " - Starting view creation..."

    ' Connect to Composer and get document
    Set swApp = New DSComposer.Application
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - No active document!"
        Exit Sub
    End If

    ' *** Processing Parts with AssySelectionMode = 0 ***
    LogPrint logFile, Now & " - Processing Parts (non-duplicates only) "
    Dim processedParts As New Collection ' Tracks unique part numbers
    Dim partCount As Long
    Dim actorSelection As String
    partCount = 0

    ' Get all actors for parts
    swDoc.AssySelectionMode = 0
    actorSelection = swDoc.GetAllActors(0)
    LogPrint logFile, Now & " - Raw actorSelection for parts (length " & Len(actorSelection) & "): '" & actorSelection & "'"
    If Len(actorSelection) > 0 Then
        actors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
        If UBound(actors) >= 0 Then
            LogPrint logFile, Now & " - Found " & UBound(actors) + 1 & " actors for parts"
            For i = 0 To UBound(actors)
                actors(i) = Trim(Replace(actors(i), "<CLitSelection Name="""" Ident=""-1"">", ""))
                LogPrint logFile, Now & " - Part actor " & i & ": '" & actors(i) & "'"
            Next i
        Else
            LogPrint logFile, Now & " - No part actors found!"
            GoTo Cleanup
        End If
    Else
        LogPrint logFile, Now & " - No part actors found!"
        GoTo Cleanup
    End If

    ' Process parts with non-numeric or null Meta.PartNo
    For i = 0 To UBound(actors)
        On Error Resume Next
        actorPath = Trim(actors(i))
        LogPrint logFile, Now & " - Checking part actor " & i & ": '" & actorPath & "'"
        
        If Len(actorPath) > 0 Then
            ' Check if actor is a part
            Dim isPart As Boolean
            isPart = False
            segments = Split(actorPath, ".")
            If UBound(segments) >= 0 Then
                Dim lastSegment As String
                lastSegment = segments(UBound(segments))
                LogPrint logFile, Now & " - Last segment for actor " & i & ": '" & lastSegment & "'"
                If Left(lastSegment, 1) = "P" Then
                    isPart = True
                End If
            End If
            
            If isPart Then
                ' Check Meta.PartNo
                Dim selectionString As String
                selectionString = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"
                Dim partPartNo As String
                partPartNo = swDoc.GetPropertyValue(selectionString, "Meta.PartNo")
                Dim cleanPartPartNo As String
                cleanPartPartNo = StripMetaPartNo(partPartNo)
                LogPrint logFile, Now & " - Meta.PartNo for part: '" & actorPath & "': '" & cleanPartPartNo & "'"
                
                ' Check if Meta.PartNo is numeric or null
                Dim isNumericPartNo As Boolean
                isNumericPartNo = True
                If Len(cleanPartPartNo) > 0 Then
                    For j = 1 To Len(cleanPartPartNo)
                        char = Mid(cleanPartPartNo, j, 1)
                        If Not IsNumeric(char) Then
                            isNumericPartNo = False
                            Exit For
                        End If
                    Next j
                Else
                    isNumericPartNo = False ' Null Meta.PartNo treated as non-numeric
                End If
                LogPrint logFile, Now & " - Is Meta.PartNo numeric for part: '" & actorPath & "'? " & isNumericPartNo
                
                If Not isNumericPartNo Then
                    LogPrint logFile, Now & " - Processing part actor: '" & actorPath & "'"
                    partCount = partCount + 1
                    LogPrint logFile, Now & " - Part count: " & partCount
                    
                    ' Select the part
                    If swDoc.SelectOnFilter(selectionString, False) Then
                        LogPrint logFile, Now & " - Part selected: '" & actorPath & "'"
                        
                        ' Try Meta.PARTNO_cfg first
                        Dim partNoCfg As String
                        partNoCfg = swDoc.GetPropertyValue(selectionString, "Meta.PARTNO_cfg")
                        Dim cleanPartNoCfg As String
                        cleanPartNoCfg = StripMetaPartNo(partNoCfg)
                        LogPrint logFile, Now & " - Meta.PARTNO_cfg: '" & cleanPartNoCfg & "'"
                        
                        ' Try Meta.PartNo_cfg if Meta.PARTNO_cfg is empty
                        If Len(cleanPartNoCfg) = 0 Then
                            Dim partNoCfgAlt As String
                            partNoCfgAlt = swDoc.GetPropertyValue(selectionString, "Meta.PartNo_cfg")
                            Dim cleanPartNoCfgAlt As String
                            cleanPartNoCfgAlt = StripMetaPartNo(partNoCfgAlt)
                            LogPrint logFile, Now & " - Meta.PartNo_cfg: '" & cleanPartNoCfgAlt & "'"
                            
                            If Len(cleanPartNoCfgAlt) > 0 Then
                                cleanPartPartNo = cleanPartNoCfgAlt
                                LogPrint logFile, Now & " - Using Meta.PartNo_cfg: '" & cleanPartNoCfgAlt & "'"
                            Else
                                cleanPartPartNo = cleanPartPartNo ' Fallback to Meta.PartNo
                                LogPrint logFile, Now & " - No Meta.PARTNO_cfg or Meta.PartNo_cfg, using Meta.PartNo: '" & cleanPartPartNo & "'"
                            End If
                        Else
                            cleanPartPartNo = cleanPartNoCfg
                            LogPrint logFile, Now & " - Using Meta.PARTNO_cfg: '" & cleanPartNoCfg & "'"
                        End If
                        
                        ' Set view name
                        Dim partViewName As String
                        If Len(cleanPartPartNo) > 0 Then
                            partViewName = cleanPartPartNo
                        Else
                            partViewName = "Part_" & Replace(actorPath, ".", "_")
                            LogPrint logFile, Now & " - No valid part number, using fallback: '" & partViewName & "'"
                        End If
                        LogPrint logFile, Now & " - View name set to: '" & partViewName & "'"
                        
                        ' Check for duplicate
                        Dim partKey As String
                        partKey = partViewName
                        processedParts.Add partKey, partKey
                        If Err.Number = 0 Then
                            ' Create the view
                            swDoc.SetVisibility "", 0, 2
                            swDoc.SetVisibility selectionString, 1, 2
                            swDoc.ZoomFitAll
                            If swDoc.CreateView(partViewName, -1) Then
                                LogPrint logFile, Now & " - Created view: '" & partViewName & "' for " & actorPath
                            Else
                                LogPrint logFile, Now & " - Failed to create view: '" & partViewName & "' for " & actorPath
                            End If
                        Else
                            LogPrint logFile, Now & " - Skipped duplicate part: '" & actorPath & "' (View name: '" & partViewName & "')"
                            Err.Clear
                        End If
                    Else
                        LogPrint logFile, Now & " - Failed to select part: '" & actorPath & "' with selection: '" & selectionString & "'"
                    End If
                Else
                    LogPrint logFile, Now & " - Skipped part with numeric Meta.PartNo: '" & actorPath & "'"
                End If
            Else
                LogPrint logFile, Now & " - Skipped non-part actor: '" & actorPath & "'"
            End If
        Else
            LogPrint logFile, Now & " - Skipped empty actor: '" & actorPath & "'"
        End If
        On Error GoTo 0
    Next i
    
    ' Log final part count
    LogPrint logFile, Now & " - Total parts processed: " & partCount

Cleanup:
    LogPrint logFile, Now & " - CreateViews for Parts completed."
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
