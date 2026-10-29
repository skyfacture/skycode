Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Sub CreateAssemblyViews()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actors() As String
    Dim i As Long
    Dim logFile As Object, fs As Object
    Dim rootID As String
    Dim rootPartNo As String
    
    ' Set up logging
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile("C:\Temp\ComposerDebug.txt", True)
    LogPrint logFile, Now & " - Starting this root bullshit again, visibility first, you fucker..."

    ' Connect to Composer and get document
    Set swApp = New DSComposer.Application
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - No active document, you fucking moron!"
        Exit Sub
    End If

    ' Get root with CLitKAssyRoot
    Dim actorSelection As String
    actorSelection = swDoc.GetAllActorsByType("CLitKAssyRoot")
    LogPrint logFile, Now & " - Raw GetAllActorsByType('CLitKAssyRoot') output: '" & actorSelection & "'"
    If Len(actorSelection) > 0 Then
        actors = Split(actorSelection, "#")
    Else
        LogPrint logFile, Now & " - GetAllActorsByType('CLitKAssyRoot') returned fuck-all, we’re screwed!"
        Exit Sub
    End If
    
    ' *** Step 1.1: Identify the Root (Iterated - CLitKAssyRoot) ***
    If UBound(actors) >= 0 Then
        Dim actor As String
        actor = Trim(actors(0))
        If Len(actor) > 0 Then
            rootID = actor
            LogPrint logFile, Now & " - Identified root (Step 1.1): '" & rootID & "' from GetAllActorsByType"
            Dim parentSelection As String
            parentSelection = swDoc.GetAssyParents(rootID)
            LogPrint logFile, Now & " - GetAssyParents of rootID: '" & parentSelection & "'"
            If Len(parentSelection) = 0 Or parentSelection = "<CLitSelection Name="" Ident=""-1""></CLitSelection>" Then
                LogPrint logFile, Now & " - Confirmed: '" & rootID & "' is the root, no parents!"
            Else
                LogPrint logFile, Now & " - Fuck, it’s a child! Parent: '" & parentSelection & "'"
                rootID = parentSelection
            End If
        Else
            LogPrint logFile, Now & " - First actor is empty, what a fucking disaster!"
            Exit Sub
        End If
    Else
        LogPrint logFile, Now & " - No actors found with CLitKAssyRoot type, this is bullshit!"
        Exit Sub
    End If

    ' *** Step 3.1: Process Root (Iterated - Show First) ***
    LogPrint logFile, Now & " - Processing root with ID: '" & rootID & "'"
    Dim selectionString As String
    selectionString = rootID
    LogPrint logFile, Now & " - Selection string: '" & selectionString & "'"
    If swDoc.SelectOnFilter(selectionString, False) Then
        LogPrint logFile, Now & " - Root selected successfully, you goddamn genius!"
        
        ' Show the root first - fuck hiding
        swDoc.SetVisibility selectionString, 1, 2 ' Root and kids ON
        swDoc.ZoomFitAll
        
        ' Try root properties
        rootPartNo = swDoc.GetPropertyValue(selectionString, "Meta.PartNo")
        Dim rootActorName As String
        rootActorName = swDoc.GetPropertyValue(selectionString, "Actor.Name")
        LogPrint logFile, Now & " - Root Meta.PartNo: '" & rootPartNo & "'"
        LogPrint logFile, Now & " - Root Actor.Name: '" & rootActorName & "'"
        
        ' If blank, get assemblies with AssySelectionMode
        If Len(rootPartNo) = 0 Then
            swDoc.AssySelectionMode = 1 ' Assemblies only
            Dim allActors As String
            allActors = swDoc.GetAllActors(1)
            LogPrint logFile, Now & " - GetAllActors(1) with AssySelectionMode: '" & allActors & "'"
            If Len(allActors) > 0 Then
                Dim subActors() As String
                subActors = Split(Replace(Replace(allActors, "<CLitSelection Name="" Ident=""-1"">", ""), "</CLitSelection>", ""), "#")
                For i = 0 To UBound(subActors)
                    Dim subActor As String
                    subActor = Trim(subActors(i))
                    If Len(subActor) > 0 Then
                        Dim subSelection As String
                        subSelection = "<CLitSelection Name=""" & subActor & """>"
                        rootPartNo = swDoc.GetPropertyValue(subSelection, "Meta.PartNo")
                        LogPrint logFile, Now & " - Sub-actor '" & subActor & "' Meta.PartNo: '" & rootPartNo & "'"
                        If Len(rootPartNo) > 0 Then
                            LogPrint logFile, Now & " - Snagged Meta.PartNo: '" & rootPartNo & "' - jackpot!"
                            swDoc.SetVisibility subSelection, 1, 2 ' Show this sub-actor
                            Exit For
                        End If
                    End If
                Next i
            Else
                LogPrint logFile, Now & " - GetAllActors(1) returned fuck-all, Composer’s a prick!"
            End If
        End If
        
        ' Set view name
        Dim cleanRootPartNo As String
        cleanRootPartNo = StripMetaPartNo(rootPartNo)
        Dim rootViewName As String
        If Len(cleanRootPartNo) > 0 Then
            rootViewName = cleanRootPartNo
            LogPrint logFile, Now & " - Cleaned root part number: '" & cleanRootPartNo & "'"
        Else
            rootViewName = "Root_CLitKAssyRoot"
            LogPrint logFile, Now & " - No Meta.PartNo, using fallback: '" & rootViewName & "'"
        End If
        
        ' Create view and reinforce visibility
            swDoc.SetVisibility selectionString, 1, 2 ' Root and kids
            swDoc.SetVisibility "", 1, 1 ' All assemblies - fuck hiding
            swDoc.ZoomFitAll ' Zoom out with all the shit
            swDoc.RefreshScene 1 ' Force viewport update with crRfsScene, you clever fuck
        If swDoc.CreateView(rootViewName, -1) Then
            LogPrint logFile, Now & " - Created view: '" & rootViewName & "' - fucking finally!"
        Else
            LogPrint logFile, Now & " - Failed to create view for '" & rootViewName & "', shit’s still fucked!"
        End If
    Else
        LogPrint logFile, Now & " - Failed to select root with ID '" & rootID & "' - what the fuck now?"
    End If

    LogPrint logFile, Now & " - Done with this root crap, check the goddamn log!"
    logFile.Close
    Set logFile = Nothing
    Set fs = Nothing
    Set swDoc = Nothing
    Set swApp = Nothing
End Sub

' Helper function to log shit
Private Sub LogPrint(logFile As Object, message As String)
    logFile.WriteLine message
    ' Debug.Print message ' Uncomment for Immediate Window
End Sub

' Helper function to strip Meta.PartNo
Private Function StripMetaPartNo(partNo As String) As String
    If Len(partNo) > 0 Then
        StripMetaPartNo = Replace(partNo, """", "")
        StripMetaPartNo = Replace(StripMetaPartNo, "<", "")
        StripMetaPartNo = Replace(StripMetaPartNo, ">", "")
    Else
        StripMetaPartNo = ""
    End If
End Function
