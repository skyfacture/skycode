Attribute VB_Name = "Auto View Creation_Assy & Tree Order"
Option Compare Database
Option Explicit

Sub CreateViewsForActors()
    ' Open Immediate Window with Ctrl + G to see output
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actorSelection As String
    Dim actors() As String
    Dim actorCount As Long
    Dim i As Long
    Dim j As Long
    Dim selectResult As Boolean
    Dim orderedActors() As Variant
    Dim processLimit As Long
    
    On Error GoTo ErrorHandler
    
    ' Toggle between limited (15) and all actors
    Const LIMIT_TO_15 As Boolean = True  ' Set to False to process all actors
    processLimit = 15
    
    Debug.Print "Starting view creation process..."
    Set swApp = New DSComposer.Application
    If swApp Is Nothing Then
        Debug.Print "Failed to connect to Composer!"
        Exit Sub
    End If
    Debug.Print "Connected to Composer!"
    
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        Debug.Print "No active document found - ensure an .smg file is open in Composer."
        Exit Sub
    End If
    Debug.Print "GetActiveDocument worked!"
    Debug.Print "FileName: " & swDoc.FileName
    
    ' Get all actors
    swDoc.AssySelectionMode = 0  ' Default to parts; adjust if assemblies are missing
    actorSelection = swDoc.GetAllActors(1) ' crAtvVisible = 1
    If Len(actorSelection) = 0 Then
        Debug.Print "No actors found."
        Exit Sub
    End If
    Debug.Print "Actor Selection: " & Left(actorSelection, 100) & "..."
    
    ' Parse actors and filter empty entries
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
    j = 0
    For i = 0 To UBound(rawActors)
        If Len(Trim(rawActors(i))) > 0 Then
            actors(j) = rawActors(i)
            j = j + 1
        End If
    Next i
    actorCount = validCount
    Debug.Print "Found " & actorCount & " valid actors."
    
    ' Build ordered list with Assy.NodeType
    ReDim orderedActors(actorCount - 1)
    For i = 0 To actorCount - 1
        Dim nodeType As String
        nodeType = swDoc.GetPropertyValue("<CLitSelection Name="""">" & actors(i) & "</CLitSelection>", "Assy.NodeType")
        Dim nodeValue As Integer
        nodeValue = Val(StripNodeType(nodeType))
        orderedActors(i) = Array(nodeValue, actors(i))
    Next i
    
    ' Sort by NodeType (descending: assemblies first)
    For i = 0 To actorCount - 2
        For j = i + 1 To actorCount - 1
            If orderedActors(i)(0) < orderedActors(j)(0) Then
                Dim temp As Variant
                temp = orderedActors(i)
                orderedActors(i) = orderedActors(j)
                orderedActors(j) = temp
            End If
        Next j
    Next i
    
    ' List ordered actors (up to limit or all)
    Debug.Print "Actors in assembly tree order:"
    Dim listLimit As Long
    listLimit = IIf(LIMIT_TO_15 And actorCount > processLimit, processLimit, actorCount)
    For i = 0 To listLimit - 1
        Debug.Print "Actor " & i & " ID: " & orderedActors(i)(1) & " (NodeType: " & orderedActors(i)(0) & ")"
    Next i
    
    ' Process actors up to limit or all
    Dim processCount As Long
    processCount = IIf(LIMIT_TO_15 And actorCount > processLimit, processLimit, actorCount)
    For i = 0 To processCount - 1
        Debug.Print "Processing actor " & i & ": " & orderedActors(i)(1)
        
        ' Select the actor
        selectResult = swDoc.SelectOnFilter("<CLitSelection Name="""">" & orderedActors(i)(1) & "</CLitSelection>", False)
        If selectResult Then
            Debug.Print "Selected actor " & i & " successfully."
        Else
            Debug.Print "Failed to select actor " & i & " with SelectOnFilter."
            GoTo NextActor
        End If
      
        ' Hide all actors
        swDoc.SetVisibility "", 0, 2
        
        ' Show only this actor
        swDoc.SetVisibility "<CLitSelection Name="""">" & orderedActors(i)(1) & "</CLitSelection>", 1, 2
        
        ' Zoom to fit all
        swDoc.ZoomFitAll
        If Err.Number <> 0 Then
            Debug.Print "Zoom failed for actor " & i & ": " & Err.Description
            Err.Clear
            GoTo NextActor
        End If
        Debug.Print "Zoomed to actor " & i & "."
        
        ' Create view with simple numbered name
        Dim initialViewName As String
        initialViewName = "View " & (i + 1)
        If swDoc.CreateView(initialViewName, i) Then
            Debug.Print "Created view '" & initialViewName & "' at position " & i & "."
            
            ' Get the Meta.PartNo property value
            Dim partNo As String
            partNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & orderedActors(i)(1) & "</CLitSelection>", "Meta.PartNo")
            Debug.Print "Meta.PartNo for actor " & i & ": " & partNo
            
            ' Rename the view using stripped Meta.PartNo value
            Dim newViewName As String
            If Len(partNo) > 0 And partNo <> "Meta.PartNo" Then
                newViewName = StripMetaPartNo(partNo)
                If Len(newViewName) > 0 Then
                    If swDoc.RenameView(initialViewName, newViewName) Then
                        Debug.Print "Renamed '" & initialViewName & "' to '" & newViewName & "'."
                    Else
                        Debug.Print "Failed to rename '" & initialViewName & "' to '" & newViewName & "'."
                    End If
                Else
                    Debug.Print "Failed to extract valid Meta.PartNo value for actor " & i & "; keeping '" & initialViewName & "'."
                End If
            Else
                Debug.Print "No valid Meta.PartNo found for actor " & i & "; keeping '" & initialViewName & "'."
            End If
        Else
            Debug.Print "Failed to create view '" & initialViewName & "'."
        End If
        
NextActor:
    Next i
    
    Debug.Print "View creation process completed."
    Exit Sub

ErrorHandler:
    Debug.Print "Error occurred: " & Err.Number & " - " & Err.Description
    On Error GoTo 0
End Sub

' Helper function to strip XML tags from Meta.PartNo
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
        strippedValue = Trim(Replace(Replace(partNo, "<Meta.PartNo>", ""), "</Meta.PartNo>", ""))
        If Len(strippedValue) > 0 Then
            StripMetaPartNo = strippedValue
        Else
            StripMetaPartNo = ""
        End If
    End If
End Function

' Helper function to strip Assy.NodeType value
Function StripNodeType(nodeType As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim strippedValue As String
    
    startPos = InStr(nodeType, "Value=""") + 7
    endPos = InStr(startPos, nodeType, """")
    If startPos > 7 And endPos > startPos Then
        strippedValue = Mid(nodeType, startPos, endPos - startPos)
        StripNodeType = strippedValue
    Else
        StripNodeType = "0"  ' Default to 0 if parsing fails
    End If
End Function
