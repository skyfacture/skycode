Attribute VB_Name = "IdnetifyRoot"
Option Compare Database
Option Explicit

' ============================================================
' Skycode - Identify Root
' ============================================================

Sub IdentifyRoot()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actorSelection As String
    Dim actors() As String
    Dim actorCount As Long
    Dim i As Long
    Dim selectResult As Boolean
    Dim processedPartNos As Collection
    Dim fs As Object, logFile As Object
    
    On Error GoTo ErrorHandler
    
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
    
    Set processedPartNos = New Collection
    
    For i = 0 To actorCount - 1
        Dim cleanActorID As String
        cleanActorID = Replace(actors(i), "<CLitSelection Name="" Ident=""-1"">", "")
        Dim topLevelID As String
        topLevelID = Left(cleanActorID, InStr(cleanActorID & ".", ".") - 1)
        topLevelID = Replace(topLevelID, "<CLitSelection Name="" Ident=""-1"">", "") ' Fix cleaning
        LogPrint logFile, "Actor " & i & " Raw ID: " & actors(i)
        LogPrint logFile, "Actor " & i & " Clean ID: " & cleanActorID
        LogPrint logFile, "Actor " & i & " Top-Level ID: " & topLevelID
        
        Dim partNo As String
        partNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", "Meta.PartNo")
        Dim cleanPartNo As String
        cleanPartNo = StripMetaPartNo(partNo) ' No Trim?preserve spaces
        LogPrint logFile, "Meta.PartNo (Top-Level ID): '" & cleanPartNo & "'"
        
        Dim actorName As String
        actorName = swDoc.GetPropertyValue("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", "Actor.Name")
        Dim cleanActorName As String
        cleanActorName = StripMetaPartNo(actorName)
        LogPrint logFile, "Actor.Name (Top-Level ID): '" & cleanActorName & "'"
        
        ' Filter by BOM top-level part numbers or ID
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
            GoTo NextActor
        End If
        
        On Error Resume Next
        processedPartNos.Add cleanPartNo, cleanPartNo
        If Err.Number = 0 Then
            LogPrint logFile, "Processing assembly " & i & ": " & cleanActorID & " (Meta.PartNo: '" & cleanPartNo & "')"
        Else
            LogPrint logFile, "Skipping duplicate assembly " & i & ": " & cleanActorID & " (Meta.PartNo: '" & cleanPartNo & "')"
            Err.Clear
            GoTo NextActor
        End If
        On Error GoTo ErrorHandler
        
        ' Select top-level ID for assemblies
        selectResult = swDoc.SelectOnFilter("<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", False)
        If selectResult Then
            LogPrint logFile, "Selected assembly " & i & " successfully with Top-Level ID: " & topLevelID
        Else
            LogPrint logFile, "Failed to select assembly " & i & " with SelectOnFilter on Top-Level ID: " & topLevelID
            processedPartNos.Remove cleanPartNo
            GoTo NextActor
        End If
      
        swDoc.SetVisibility "", 0, 2
        swDoc.SetVisibility "<CLitSelection Name="""">" & topLevelID & "</CLitSelection>", 1, 2
        
        swDoc.ZoomFitAll
        If Err.Number <> 0 Then
            LogPrint logFile, "Zoom failed for assembly " & i & ": " & Err.description
            Err.Clear
            processedPartNos.Remove cleanPartNo
            GoTo NextActor
        End If
        LogPrint logFile, "Zoomed to assembly " & i & "."
        
        Dim initialViewName As String
        initialViewName = "View " & (i + 1)
        If swDoc.CreateView(initialViewName, i) Then
            LogPrint logFile, "Created view '" & initialViewName & "' at position " & i & "."
            
            Dim newViewName As String
            If Len(cleanPartNo) > 0 Then
                newViewName = cleanPartNo
                If swDoc.RenameView(initialViewName, newViewName) Then
                    LogPrint logFile, "Renamed '" & initialViewName & "' to '" & newViewName & "'."
                Else
                    LogPrint logFile, "Failed to rename '" & initialViewName & "' to '" & newViewName & "'."
                    processedPartNos.Remove cleanPartNo
                End If
            Else
                LogPrint logFile, "No valid Meta.PartNo for assembly " & i & "; keeping '" & initialViewName & "'."
            End If
        Else
            LogPrint logFile, "Failed to create view '" & initialViewName & "'."
            processedPartNos.Remove cleanPartNo
        End If
        
NextActor:
    Next i
    
    If processedPartNos.Count >= 4 Then
        LogPrint logFile, "Found all assemblies/sub-assemblies."
    Else
        LogPrint logFile, "Warning: Only found " & processedPartNos.Count & " unique assemblies/sub-assemblies."
    End If
    
    LogPrint logFile, "Assembly view creation process completed."
    If Not logFile Is Nothing Then
        logFile.Close
    End If
    Set logFile = Nothing
    Set fs = Nothing
    Exit Sub

ErrorHandler:
    If Not logFile Is Nothing Then
        LogPrint logFile, "Error occurred: " & Err.Number & " - " & Err.description
        logFile.Close
    Else
        Debug.Print "Error occurred before log file creation: " & Err.Number & " - " & Err.description
    End If
    Set logFile = Nothing
    Set fs = Nothing
    On Error GoTo 0
End Sub

Sub LogPrint(logFile As Object, message As String)
    Debug.Print message
    If Not logFile Is Nothing Then
        logFile.WriteLine Now & " - " & message
    Else
        Debug.Print "Log file not initialized: " & message
    End If
End Sub

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




