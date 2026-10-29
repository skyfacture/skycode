Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Sub TestPartPropertyFix()
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
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\Temp\PartFixDebug.txt", True)
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

    ' Get all actors
    swDoc.AssySelectionMode = 0 ' All actors to ensure we get parts
    actorSelection = swDoc.GetAllActors(0)
    LogPrint logFile, Now & " - Raw actorSelection (length " & Len(actorSelection) & "): '" & actorSelection & "'"
    If Len(actorSelection) > 0 Then
        actors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
        If UBound(actors) >= 0 Then
            LogPrint logFile, Now & " - Found " & UBound(actors) + 1 & " actors after split"
            ' Log all actors
            For i = 0 To UBound(actors)
                LogPrint logFile, Now & " - Actor " & i & ": '" & Trim(actors(i)) & "'"
            Next i
        Else
            LogPrint logFile, Now & " - No valid actors after split!"
            Exit Sub
        End If
    Else
        LogPrint logFile, Now & " - No actors found!"
        Exit Sub
    End If
    
    ' Process parts with non-numeric Meta.PartNo
    partCount = 0
    For i = 0 To UBound(actors)
        On Error Resume Next
        actorPath = Trim(actors(i))
        ' Clean up any leading CLitSelection tags
        If InStr(actorPath, ">") > 0 Then
            actorPath = Mid(actorPath, InStr(actorPath, ">") + 1)
        End If
        actorPath = Trim(actorPath)
        LogPrint logFile, Now & " - Checking actor " & i & ": '" & actorPath & "'"
        
        ' Check if actor is a part
        Dim isPart As Boolean
        isPart = False
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            If UBound(segments) >= 0 Then
                Dim lastSegment As String
                lastSegment = segments(UBound(segments))
                LogPrint logFile, Now & " - Last segment for actor " & i & ": '" & lastSegment & "'"
                If Left(lastSegment, 1) = "P" Then
                    isPart = True
                End If
            End If
        End If
        
        If isPart Then
            ' Check Meta.PartNo before processing
            Dim selectionString As String
            selectionString = "<CLitSelection Name="""">" & actorPath & "</CLitSelection>"
            partPartNo = swDoc.GetPropertyValue(selectionString, "Meta.PartNo")
            cleanPartPartNo = StripMetaPartNo(partPartNo)
            LogPrint logFile, Now & " - Meta.PartNo for actor " & i & ": '" & cleanPartPartNo & "'"
            
            ' Check if Meta.PartNo is numeric
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
                isNumericPartNo = False
            End If
            LogPrint logFile, Now & " - Is Meta.PartNo numeric for actor " & i & "? " & isNumericPartNo
            
            If Not isNumericPartNo Then
                LogPrint logFile, Now & " - Processing part actor: '" & actorPath & "'"
                partCount = partCount + 1
                LogPrint logFile, Now & " - Part count: " & partCount
                
                ' Select the part
                If swDoc.SelectOnFilter(selectionString, False) Then
                    LogPrint logFile, Now & " - Part selected: '" & actorPath & "'"
                    
                    ' Use Meta.PARTNO_cfg
                    partNoCfg = swDoc.GetPropertyValue(selectionString, "Meta.PARTNO_cfg")
                    cleanPartNoCfg = StripMetaPartNo(partNoCfg)
                    LogPrint logFile, Now & " - Meta.PARTNO_cfg: '" & cleanPartNoCfg & "'"
                    
                    If Len(cleanPartNoCfg) > 0 Then
                        cleanPartPartNo = cleanPartNoCfg
                        LogPrint logFile, Now & " - Using Meta.PARTNO_cfg: '" & cleanPartNoCfg & "'"
                    Else
                        cleanPartPartNo = cleanPartPartNo ' Fallback to Meta.PartNo if PARTNO_cfg empty
                        LogPrint logFile, Now & " - No Meta.PARTNO_cfg, using Meta.PartNo: '" & cleanPartPartNo & "'"
                    End If
                    
                    ' Set view name
                    If Len(cleanPartPartNo) > 0 Then
                        partViewName = cleanPartPartNo
                    Else
                        partViewName = "Part_" & Replace(actorPath, ".", "_")
                        LogPrint logFile, Now & " - No valid part number, using fallback: '" & partViewName & "'"
                    End If
                    LogPrint logFile, Now & " - View name set to: '" & partViewName & "'"
                    
                    ' Check for duplicate part number
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
            StripMetaPartNo = Trim(Mid(partNo, startPos, endPos - startPos))
            Exit Function
        End If
    End If
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
