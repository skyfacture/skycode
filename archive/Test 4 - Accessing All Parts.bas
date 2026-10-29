Attribute VB_Name = "Auto View Creation"
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
    Dim selectResult As Boolean
    
    On Error GoTo ErrorHandler
    
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
    
    actorSelection = swDoc.GetAllActors(1) ' crAtvVisible = 1
    If Len(actorSelection) = 0 Then
        Debug.Print "GetAllActors returned no actors."
        Exit Sub
    End If
    Debug.Print "Actor Selection: " & Left(actorSelection, 100) & "..."
    
    ' Parse actors by removing XML tags and splitting on #
    actors = Split(Replace(Replace(actorSelection, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
    actorCount = UBound(actors) + 1
    Debug.Print "Found " & actorCount & " actors."
    ' Limit listing to first 5 actors or total count if less than 5
    Dim listCount As Long
    listCount = IIf(actorCount > 5, 5, actorCount)
    For i = 0 To listCount - 1
        Debug.Print "Actor " & i & " ID: " & actors(i)
    Next i
    
    ' Process only the first 5 actors
    For i = 0 To IIf(actorCount > 5, 4, actorCount - 1)
        Debug.Print "Processing actor " & i & ": " & actors(i)
        
        ' Select the actor
        selectResult = swDoc.SelectOnFilter("<CLitSelection Name="""">" & actors(i) & "</CLitSelection>", False)
        If selectResult Then
            Debug.Print "Selected actor " & i & " successfully."
        Else
            Debug.Print "Failed to select actor " & i & " with SelectOnFilter."
            GoTo NextActor
        End If
      
        ' Hide all actors
        swDoc.SetVisibility "", 0, 2  ' crAtvInvisible = 2 for unspecified actors
        
        ' Show only this actor
        swDoc.SetVisibility "<CLitSelection Name="""">" & actors(i) & "</CLitSelection>", 1, 2  ' crAtvVisible = 1
        
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
        initialViewName = "View " & (i + 1)  ' e.g., "View 1", "View 2", etc.
        If swDoc.CreateView(initialViewName, i) Then
            Debug.Print "Created view '" & initialViewName & "' at position " & i & "."
            
            ' Get the Meta.PartNo property value
            Dim partNo As String
            partNo = swDoc.GetPropertyValue("<CLitSelection Name="""">" & actors(i) & "</CLitSelection>", "Meta.PartNo")
            Debug.Print "Meta.PartNo for actor " & i & ": " & partNo
            
            ' Rename the view using stripped Meta.PartNo value
            Dim newViewName As String
            If Len(partNo) > 0 And partNo <> "Meta.PartNo" Then
                newViewName = StripMetaPartNo(partNo)  ' Extract the raw value
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
    
    ' Look for Value="..." pattern
    startPos = InStr(partNo, "Value=""") + 7  ' Move past 'Value="'
    endPos = InStr(startPos, partNo, """")   ' Find closing quote
    If startPos > 7 And endPos > startPos Then
        strippedValue = Mid(partNo, startPos, endPos - startPos)
        StripMetaPartNo = strippedValue
    Else
        ' Fallback: if no Value="..." found, return original string trimmed
        strippedValue = Trim(Replace(Replace(partNo, "<Meta.PartNo>", ""), "</Meta.PartNo>", ""))
        If Len(strippedValue) > 0 Then
            StripMetaPartNo = strippedValue
        Else
            StripMetaPartNo = ""  ' Return empty if parsing fails
        End If
    End If
End Function
