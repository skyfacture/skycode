Option Compare Database
Option Explicit
' ============================================================
' Skycode - All Views with dimensioned bounding box
' Creates views for all assemblies and non-numeric parts, including bounding box dimensions.
' Normal assembly/part passes retain the original naming and eligibility rules.
' Proxy parts bypass ordinary part naming. The final BOX view captures the root,
' its visible proxy, opacity 100 and dimensions, without updating earlier views.
' ============================================================
Private Const LOG_FILE_PATH As String = "C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\views [assys + all nonnumeric parts + bounding box dims].txt"
Private Const STYLE_PATH As String = "C:\Users\skyfa\OneDrive\Documents\SOLIDWORKSComposer\User\Profiles\Dims.smgStyleSet"
Private Const DIM_TAG As String = "SkycodeBBoxDim"
Private Const OFFSET_DISTANCE As Double = 100#
Private Const CMD_DIMENSIONS As Long = 178

Public Sub CreateViewsAllWithBoundingBoxDims()
    Dim swApp As DSComposer.Application
    Dim swDoc As DSComposer.Document
    Dim actors() As String
    Dim i As Long, j As Long
    Dim logFile As Object, fs As Object
    Dim uniqueAssemblies As Collection
    Dim rootID As String, actorPath As String, actorSelection As String
    Dim segments() As String, char As String
    Dim rootViewName As String
    Dim rootViewCreated As Boolean
    Dim previousCameraSmoothTime As Integer, cameraSmoothTimeSaved As Boolean
    Dim assemblyPath As String, assemblyID As Variant
    Dim subPartNo As String, cleanSubPartNo As String, subViewName As String
    Dim processedParts As New Collection
    Dim partCount As Long, isPart As Boolean, lastSegment As String
    Dim selectionString As String, partPartNo As String, cleanPartPartNo As String
    Dim isNumericPartNo As Boolean
    Dim partNoCfg As String, cleanPartNoCfg As String
    Dim partNoCfgAlt As String, cleanPartNoCfgAlt As String
    Dim partViewName As String, partKey As String, actorName As String
    Dim errorNumber As Long, errorDescription As String

    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile(LOG_FILE_PATH, True)
    LogPrint logFile, Now & " - Starting combined view creation..."
    Set swApp = New DSComposer.Application
    Set swDoc = swApp.GetActiveDocument
    If swDoc Is Nothing Then
        LogPrint logFile, Now & " - ERROR: No active document!"
        GoTo Cleanup
    End If

    previousCameraSmoothTime = swDoc.CameraSmoothTime
    cameraSmoothTimeSaved = True
    swDoc.CameraSmoothTime = 0
    swDoc.AssySelectionMode = 1
    rootID = "-1"
    rootViewName = ReadRootActorName(swDoc, rootID)
    LogPrint logFile, Now & " - Document root: " & rootID & "; Actor.Name: " & rootViewName
    actorSelection = swDoc.GetAllActors(1)
    actors = SelectionIds(actorSelection)
    Set uniqueAssemblies = New Collection
    For i = 0 To UBound(actors)
        actorPath = Trim(actors(i))
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            assemblyPath = ""
            For j = 0 To UBound(segments) - 1
                If assemblyPath = "" Then
                    assemblyPath = segments(j)
                Else
                    assemblyPath = assemblyPath & "." & segments(j)
                End If
                AddUniqueAssembly uniqueAssemblies, assemblyPath
            Next j
        End If
    Next i

    If Len(rootID) > 0 And Len(rootViewName) > 0 Then
        If swDoc.SelectOnFilter(Sel(rootID), False) Then
            swDoc.SetVisibility "", 0, 2
            swDoc.SetVisibility Sel(rootID), 1, 2
            HideCollaborations swDoc
            swDoc.ZoomFitAll
            rootViewCreated = swDoc.CreateView(rootViewName, -1)
            If rootViewCreated Then
                LogPrint logFile, Now & " - Created view: " & rootViewName & " for " & rootID
            Else
                LogPrint logFile, Now & " - ERROR: Failed to create root view: " & rootViewName
            End If
        Else
            LogPrint logFile, Now & " - Failed to select top sub-assy: " & rootID
        End If
    Else
        LogPrint logFile, Now & " - ERROR: Root ID or Actor.Name is missing; root and BOX views will not be created."
    End If

    For Each assemblyID In uniqueAssemblies
        If assemblyID <> rootID Then
            If swDoc.SelectOnFilter(Sel(CStr(assemblyID)), False) Then
                swDoc.SetVisibility "", 0, 2
                swDoc.SetVisibility Sel(CStr(assemblyID)), 1, 2
                HideCollaborations swDoc
                swDoc.ZoomFitAll
                subPartNo = swDoc.GetPropertyValue(Sel(CStr(assemblyID)), "Meta.PartNo")
                cleanSubPartNo = StripMetaPartNo(subPartNo)
                If Len(cleanSubPartNo) > 0 Then
                    subViewName = cleanSubPartNo
                Else
                    subViewName = "SubAssy_" & Replace(assemblyID, ".", "_")
                End If
                If swDoc.CreateView(subViewName, -1) Then
                    LogPrint logFile, Now & " - Created view: " & subViewName & " for " & assemblyID
                Else
                    LogPrint logFile, Now & " - Failed to create view: " & subViewName
                End If
            Else
                LogPrint logFile, Now & " - Failed to select sub-assembly: " & assemblyID
            End If
        End If
    Next assemblyID

    swDoc.AssySelectionMode = 0
    actorSelection = swDoc.GetAllActors(0)
    actors = SelectionIds(actorSelection)
    For i = 0 To UBound(actors)
        On Error GoTo PartFailed
        actorPath = Trim(actors(i))
        If Len(actorPath) > 0 Then
            segments = Split(actorPath, ".")
            lastSegment = segments(UBound(segments))
            isPart = (Left(lastSegment, 1) = "P")
            If isPart Then
                selectionString = Sel(actorPath)
                actorName = StripMetaPartNo(swDoc.GetPropertyValue(selectionString, "Actor.Name"))
                If IsBoxProxy(actorName) Then
                    LogPrint logFile, Now & " - Reserved proxy for BOX phase: " & actorPath
                    GoTo NextPart
                End If
                partPartNo = swDoc.GetPropertyValue(selectionString, "Meta.PartNo")
                cleanPartPartNo = StripMetaPartNo(partPartNo)
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
                If Not isNumericPartNo Then
                    partCount = partCount + 1
                    If swDoc.SelectOnFilter(selectionString, False) Then
                        partNoCfg = swDoc.GetPropertyValue(selectionString, "Meta.PARTNO_cfg")
                        cleanPartNoCfg = StripMetaPartNo(partNoCfg)
                        If Len(cleanPartNoCfg) = 0 Then
                            partNoCfgAlt = swDoc.GetPropertyValue(selectionString, "Meta.PartNo_cfg")
                            cleanPartNoCfgAlt = StripMetaPartNo(partNoCfgAlt)
                            If Len(cleanPartNoCfgAlt) > 0 Then cleanPartPartNo = cleanPartNoCfgAlt
                        Else
                            cleanPartPartNo = cleanPartNoCfg
                        End If
                        If Len(cleanPartPartNo) > 0 Then
                            partViewName = cleanPartPartNo
                        Else
                            partViewName = "Part_" & Replace(actorPath, ".", "_")
                        End If
                        partKey = partViewName
                        If AddUniquePart(processedParts, partKey) Then
                            swDoc.SetVisibility "", 0, 2
                            swDoc.SetVisibility selectionString, 1, 2
                            HideCollaborations swDoc
                            swDoc.ZoomFitAll
                            If swDoc.CreateView(partViewName, -1) Then
                                LogPrint logFile, Now & " - Created view: '" & partViewName & "' for " & actorPath
                            Else
                                LogPrint logFile, Now & " - Failed to create view: '" & partViewName & "' for " & actorPath
                            End If
                        Else
                            LogPrint logFile, Now & " - Skipped duplicate part: '" & actorPath & "' (View name: '" & partViewName & "')"
                        End If
                    Else
                        LogPrint logFile, Now & " - Failed to select part: " & actorPath
                    End If
                Else
                    LogPrint logFile, Now & " - Skipped part with numeric Meta.PartNo: " & actorPath
                End If
            End If
        End If
NextPart:
        On Error GoTo Failed
    Next i
    LogPrint logFile, Now & " - Total parts processed: " & partCount

    If rootViewCreated Then
        CreateBoxView swApp, swDoc, rootID, rootViewName, logFile
    Else
        LogPrint logFile, Now & " - ERROR: BOX phase skipped because the root view was not created."
    End If
    LogPrint logFile, Now & " - Combined view creation completed."
    GoTo Cleanup

PartFailed:
    LogPrint logFile, Now & " - ERROR: Part processing failed for " & actorPath & ": " & Err.Number & ": " & Err.Description
    Resume NextPart
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    LogPrint logFile, Now & " - ERROR " & errorNumber & ": " & errorDescription
    MsgBox "Combined view creation failed:" & vbCrLf & errorDescription, vbExclamation
Cleanup:
    If cameraSmoothTimeSaved Then RestoreCameraSmoothTime swDoc, previousCameraSmoothTime, logFile
    If Not logFile Is Nothing Then logFile.Close
End Sub

Private Function ReadRootActorName(ByVal doc As DSComposer.Document, ByVal rootID As String) As String
    Dim xml As Object, nameNode As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    If Not xml.LoadXML(doc.GetPropertyValue(Sel(rootID), "Actor.Name")) Then
        Err.Raise vbObjectError + 3104, , "Invalid Actor.Name XML for root " & rootID & ": " & xml.parseError.reason
    End If
    Set nameNode = xml.SelectSingleNode("//Actor.Name/@Value")
    If Not nameNode Is Nothing Then ReadRootActorName = Trim$(nameNode.Text)
End Function

Private Sub RestoreCameraSmoothTime(ByVal doc As DSComposer.Document, ByVal previousValue As Integer, _
    ByVal logFile As Object)
    On Error GoTo Failed
    doc.CameraSmoothTime = previousValue
    Exit Sub
Failed:
    LogPrint logFile, Now & " - ERROR: Could not restore CameraSmoothTime: " & Err.Number & ": " & Err.Description
    MsgBox "Could not restore CameraSmoothTime to " & previousValue & ":" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub AddUniqueAssembly(ByVal assemblies As Collection, ByVal path As String)
    On Error GoTo Failed
    assemblies.Add path, path
    Exit Sub
Failed:
    If Err.Number <> 457 Then Err.Raise Err.Number, "AddUniqueAssembly", Err.Description
End Sub

Private Function AddUniquePart(ByVal parts As Collection, ByVal key As String) As Boolean
    On Error GoTo Failed
    parts.Add key, key
    AddUniquePart = True
    Exit Function
Failed:
    If Err.Number <> 457 Then Err.Raise Err.Number, "AddUniquePart", Err.Description
End Function

Private Function IsBoxProxy(ByVal actorName As String) As Boolean
    IsBoxProxy = (InStr(1, actorName, "BoundingBox", vbTextCompare) > 0 Or _
        InStr(1, actorName, "BestFitEnvelope", vbTextCompare) > 0) And _
        InStr(1, actorName, DIM_TAG, vbBinaryCompare) = 0
End Function

Private Sub HideCollaborations(ByVal doc As DSComposer.Document, Optional ByVal keepBoxDimensions As Boolean = False)
    Dim ids As Variant, id As Variant, actorName As String
    doc.GroundGrid = False
    ids = SelectionIds(doc.GetAllCollaborations)
    For Each id In ids
        If Len(CStr(id)) > 0 Then
            If keepBoxDimensions Then
                actorName = doc.GetPropertyValue(Sel(CStr(id)), "Actor.Name")
                If InStr(1, actorName, DIM_TAG & " ", vbBinaryCompare) > 0 Then GoTo NextCollaboration
            End If
            doc.SetVisibility Sel(CStr(id)), 2, 0
            If doc.GetVisibility(Sel(CStr(id))) <> 2 Then
                Err.Raise vbObjectError + 3100, , "Could not hide collaboration actor: " & id
            End If
        End If
NextCollaboration:
    Next id
End Sub

Private Sub CreateBoxView(ByVal app As DSComposer.Application, ByVal doc As DSComposer.Document, _
    ByVal rootID As String, ByVal rootViewName As String, ByVal logFile As Object)

    Dim ids As Variant, id As Variant, actorName As String, styleProps As String
    Dim candidates As New Collection, candidate As Variant
    Dim origin As Variant, edge As Variant, i As Long
    Dim boxViewName As String, oldSelection As String

    On Error GoTo Failed
    If Len(rootViewName) = 0 Then
        LogPrint logFile, Now & " - ERROR: BOX view skipped: root has no Actor.Name."
        Exit Sub
    End If
    If Not doc.GoToView("ref:name:" & rootViewName) Then
        Err.Raise vbObjectError + 3101, , "Could not return to root view: " & rootViewName
    End If
    doc.AssySelectionMode = 0
    oldSelection = doc.Selection
    ids = SelectionIds(doc.GetAllActors(0))
    For Each id In ids
        If rootID = "-1" Or Left$(CStr(id), Len(rootID) + 1) = rootID & "." Then
            actorName = StripMetaPartNo(doc.GetPropertyValue(Sel(CStr(id)), "Actor.Name"))
            If IsBoxProxy(actorName) And doc.GetVisibility(Sel(CStr(id))) = 1 Then
                candidates.Add CStr(id)
            End If
        End If
    Next id
    If candidates.Count = 0 Then
        LogPrint logFile, Now & " - ERROR: No visible BoundingBox or BestFitEnvelope proxy under root " & rootID & ". Run Skycode - sldworks bounding box.bas and import or re-import the proxy into Composer."
        Exit Sub
    End If
    ' Preflight every candidate before deleting or creating any dimensions.
    For Each candidate In candidates
        If Not ReadVector(doc, CStr(candidate), "Meta.BestFitOrigin", origin) Then
            LogPrint logFile, Now & " - ERROR: BOX view skipped: missing Meta.BestFitOrigin for " & candidate & ". Run Skycode - sldworks bounding box.bas and re-import the proxy."
            Exit Sub
        End If
        For i = 1 To 3
            If Not ReadVector(doc, CStr(candidate), "Meta.BestFitEdge" & i, edge) Then
                LogPrint logFile, Now & " - ERROR: BOX view skipped: missing Meta.BestFitEdge" & i & " for " & candidate & ". Run Skycode - sldworks bounding box.bas and re-import the proxy."
                Exit Sub
            End If
            If Norm(edge) <= 0 Then Err.Raise vbObjectError + 3102, , "Invalid zero-length box edge for " & candidate
        Next i
    Next candidate
    styleProps = LoadStyleProperties(STYLE_PATH)
    For Each candidate In candidates
        ApplyBoundingBoxOpacity doc, CStr(candidate), logFile
    Next candidate
    HideCollaborations doc
    doc.ZoomFitAll
    For Each candidate In candidates
        If Not DimensionProxy(app, doc, CStr(candidate), styleProps, logFile) Then
            LogPrint logFile, Now & " - ERROR: BOX view not created because dimension creation failed."
            doc.Selection = oldSelection
            Exit Sub
        End If
    Next candidate
    doc.Selection = oldSelection
    boxViewName = rootViewName & " BOX"
    HideCollaborations doc, True
    If Not doc.CreateView(boxViewName, -1) Then
        Err.Raise vbObjectError + 3103, , "Could not create BOX view: " & boxViewName
    End If
    LogPrint logFile, Now & " - Created BOX view: " & boxViewName
    Exit Sub
Failed:
    LogPrint logFile, Now & " - ERROR: BOX phase failed; normal views have already been processed. " & Err.Number & ": " & Err.Description
End Sub

Private Sub ApplyBoundingBoxOpacity(ByVal doc As DSComposer.Document, ByVal actorId As String, _
    ByVal logFile As Object)
    Dim propertyMap As String
    propertyMap = "<CLitPropertyMap><CLitModifiable Ident=""" & actorId & """>" & _
        "<Actor.Alpha Value=""100""></Actor.Alpha>" & _
        "</CLitModifiable></CLitPropertyMap>"
    If Not doc.SetPropertyMap(propertyMap) Then
        Err.Raise vbObjectError + 3006, , "Could not set opacity to 100 for actor " & actorId & "."
    End If
    LogPrint logFile, Now & " - Set opacity to 100 for actor " & actorId
End Sub

Private Function DimensionProxy(ByVal app As DSComposer.Application, ByVal doc As DSComposer.Document, _
    ByVal proxyId As String, ByVal styleProps As String, ByVal logFile As Object) As Boolean

    Dim origin As Variant, edges(0 To 2) As Variant, axes(0 To 2) As Variant, lengths(0 To 2) As Double
    Dim order(0 To 2) As Long, roles As Variant, worldOffsets(0 To 2) As Variant
    Dim centre As Variant, toViewer As Variant
    Dim before As String, created As Variant, i As Long, j As Long, t As Long
    Dim e As Long, o As Long, r As Long, oSign As Double, start As Variant, offsetDir As Variant

    If Not ReadVector(doc, proxyId, "Meta.BestFitOrigin", origin) Then
        LogPrint logFile, Now & " - Skipped " & proxyId & ": no Meta.BestFitOrigin (regenerate the proxy with the SOLIDWORKS macro)."
        Exit Function
    End If
    For i = 0 To 2
        If Not ReadVector(doc, proxyId, "Meta.BestFitEdge" & (i + 1), edges(i)) Then
            LogPrint logFile, Now & " - Skipped " & proxyId & ": no Meta.BestFitEdge" & (i + 1) & "."
            Exit Function
        End If
        lengths(i) = Norm(edges(i))
        axes(i) = ScaleVec(edges(i), 1# / lengths(i))
        order(i) = i
    Next i

    ' order(0) = longest (length), order(1) = width, order(2) = shortest (height).
    For i = 0 To 1
        For j = i + 1 To 2
            If lengths(order(j)) > lengths(order(i)) Then t = order(i): order(i) = order(j): order(j) = t
        Next j
    Next i
    roles = Array("Length", "Width", "Height")
    worldOffsets(0) = V(0, -1, 0)
    worldOffsets(1) = V(0, -1, 0)
    worldOffsets(2) = V(-1, 0, 0)

    centre = Add(origin, ScaleVec(Add(Add(edges(0), edges(1)), edges(2)), 0.5))
    toViewer = Subtract(CameraPosition(doc), centre)

    RemovePreviousDimensions doc, proxyId
    before = doc.GetAllCollaborations
    doc.Selection = Sel(proxyId)
    app.RunRibbonCommand CMD_DIMENSIONS
    created = NewIds(before, doc.GetAllCollaborations)
    If UBound(created) < 2 Then
        LogPrint logFile, Now & " - Skipped " & proxyId & ": Composer created " & (UBound(created) + 1) & " dimensions, expected 3."
        For i = 0 To UBound(created): doc.Remove Sel(CStr(created(i))): Next i
        Exit Function
    End If

    For i = 0 To 2
        e = order(i)
        ' Snap the world offset to whichever of the other two box axes it matches best, keeping its sign.
        o = (e + 1) Mod 3: r = (e + 2) Mod 3
        If Abs(Dot(worldOffsets(i), axes(r))) > Abs(Dot(worldOffsets(i), axes(o))) Then t = o: o = r: r = t
        oSign = Sgn(Dot(worldOffsets(i), axes(o))): If oSign = 0 Then oSign = 1
        offsetDir = ScaleVec(axes(o), oSign)

        ' Edge on the box face in the offset direction, on the side facing the camera.
        start = origin
        If oSign > 0 Then start = Add(start, edges(o))
        If Dot(axes(r), toViewer) > 0 Then start = Add(start, edges(r))

        PlaceDimension doc, CStr(created(i)), start, Add(start, edges(e)), lengths(e), offsetDir, _
            DIM_TAG & " " & roles(i) & " " & proxyId, styleProps
        LogPrint logFile, Now & " - " & proxyId & " " & roles(i) & " " & Format$(lengths(e), "0.00") & " mm -> " & created(i)
    Next i
    For i = 3 To UBound(created): doc.Remove Sel(CStr(created(i))): Next i
    DimensionProxy = True
End Function

Private Sub PlaceDimension(ByVal doc As DSComposer.Document, ByVal id As String, a As Variant, b As Variant, _
    ByVal length As Double, offsetDir As Variant, ByVal dimName As String, ByVal styleProps As String)

    Dim frame As String, vx As Variant, vy As Variant, vz As Variant
    Dim roundedLabel As String

    ' Copy the visual properties without subscribing the dimension to the style.
    If Not doc.SetPropertyMap("<CLitPropertyMap><CLitModifiable Ident=""" & id & """>" & _
        AnchorXml(0, a) & AnchorXml(1, b) & styleProps & _
        "<Actor.Style Value=""""></Actor.Style>" & _
        "<Actor.Name Value=""" & XmlEscape(dimName) & """></Actor.Name>" & _
        "</CLitModifiable></CLitPropertyMap>") Then
        Err.Raise vbObjectError + 3108, , "Could not set unlinked dimension appearance: " & id
    End If

    ' Composer builds the dimension frame from the anchors (VX along the edge); Text.Position is local to it.
    frame = doc.GetPropertyValue(Sel(id), "Actor.Location")
    frame = Left$(frame, InStr(frame, "</Frame>"))
    vx = V(AttrValue(frame, "VX", "X"), AttrValue(frame, "VX", "Y"), AttrValue(frame, "VX", "Z"))
    vy = V(AttrValue(frame, "VY", "X"), AttrValue(frame, "VY", "Y"), AttrValue(frame, "VY", "Z"))
    vz = Cross(vx, vy)

    doc.SetPropertyMap "<CLitPropertyMap><CLitModifiable Ident=""" & id & """><Text.Position X=""" & Num(length / 2) & _
        """ Y=""" & Num(OFFSET_DISTANCE * Dot(offsetDir, vy)) & """ Z=""" & Num(OFFSET_DISTANCE * Dot(offsetDir, vz)) & _
        """></Text.Position></CLitModifiable></CLitPropertyMap>"

    ' Apply separately: the style contains conflicting automatic-value and text properties.
    roundedLabel = Format$(-Int(-length), "0")
    If Not doc.SetPropertyMap("<CLitPropertyMap><CLitModifiable Ident=""" & id & """>" & _
        "<Actor.DisplayValue Value=""0""></Actor.DisplayValue>" & _
        "<Text.MiddleString Value=""" & roundedLabel & """></Text.MiddleString>" & _
        "</CLitModifiable></CLitPropertyMap>") Then
        Err.Raise vbObjectError + 3106, , "Could not set whole-number dimension label: " & id
    End If
    If StripMetaPartNo(doc.GetPropertyValue(Sel(id), "Actor.DisplayValue")) <> "0" Or _
        StripMetaPartNo(doc.GetPropertyValue(Sel(id), "Text.MiddleString")) <> roundedLabel Then
        Err.Raise vbObjectError + 3107, , "Composer did not retain the whole-number dimension label: " & id
    End If
End Sub

Private Sub RemovePreviousDimensions(ByVal doc As DSComposer.Document, ByVal proxyId As String)
    Dim ids As Variant, id As Variant, dimName As String
    ids = SelectionIds(doc.GetAllCollaborations)
    For Each id In ids
        dimName = doc.GetPropertyValue(Sel(CStr(id)), "Actor.Name")
        If InStr(1, dimName, DIM_TAG & " ", vbBinaryCompare) > 0 And InStr(1, dimName, proxyId, vbBinaryCompare) > 0 Then
            doc.Remove Sel(CStr(id))
        End If
    Next id
End Sub

Private Function LoadStyleProperties(ByVal path As String) As String
    Dim fs As Object, text As String, re As Object, p As Long, q As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(path) Then Err.Raise vbObjectError + 3002, , "Style file not found: " & path
    text = fs.OpenTextFile(path, 1).ReadAll
    p = InStr(text, "<CLitStyleProperties>"): q = InStr(text, "</CLitStyleProperties>")
    If p = 0 Or q = 0 Then Err.Raise vbObjectError + 3003, , "Unrecognised style file: " & path
    text = Mid$(text, p + Len("<CLitStyleProperties>"), q - p - Len("<CLitStyleProperties>"))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "<(Profile\.[^ /]+|Link[^ /]*)[^>]*/>"
    text = re.Replace(text, "")
    re.Pattern = "<([A-Za-z0-9.]+)([^>]*)/>"
    LoadStyleProperties = re.Replace(text, "<$1$2></$1>")
End Function

Private Function ReadVector(ByVal doc As DSComposer.Document, ByVal id As String, ByVal prop As String, ByRef result As Variant) As Boolean
    Dim xml As String, p As Long, q As Long, parts As Variant
    xml = doc.GetPropertyValue(Sel(id), prop)
    p = InStr(xml, "Value=""")
    If p = 0 Then Exit Function
    p = p + 7: q = InStr(p, xml, """")
    parts = Split(Mid$(xml, p, q - p), ";")
    If UBound(parts) <> 2 Then Exit Function
    result = V(Val(parts(0)), Val(parts(1)), Val(parts(2)))
    ReadVector = True
End Function

Private Function CameraPosition(ByVal doc As DSComposer.Document) As Variant
    Dim cam As String
    cam = doc.GetCamera
    cam = Mid$(cam, InStr(cam, "<Navigator.Camera.Location>"))
    CameraPosition = V(AttrValue(cam, "Origin", "X"), AttrValue(cam, "Origin", "Y"), AttrValue(cam, "Origin", "Z"))
End Function

Private Function AnchorXml(ByVal index As Long, p As Variant) As String
    AnchorXml = "<Anchor." & index & " AnchorId=""" & index & """ DragAble=""0""><Frame><Origin X=""" & Num(p(0)) & _
        """ Y=""" & Num(p(1)) & """ Z=""" & Num(p(2)) & """></Origin><VX X=""1"" Y=""0"" Z=""0""></VX>" & _
        "<VY X=""0"" Y=""1"" Z=""0""></VY></Frame></Anchor." & index & ">"
End Function

Private Function AttrValue(ByVal xml As String, ByVal tag As String, ByVal attr As String) As Double
    Dim p As Long, q As Long
    p = InStr(xml, "<" & tag & " ")
    If p = 0 Then Err.Raise vbObjectError + 3004, , "Missing <" & tag & "> in Composer response."
    q = InStr(p, xml, ">")
    p = InStr(p, xml, " " & attr & "=""")
    If p = 0 Or p > q Then Err.Raise vbObjectError + 3005, , "Missing " & tag & "." & attr & " in Composer response."
    p = p + Len(attr) + 3
    AttrValue = Val(Mid$(xml, p, InStr(p, xml, """") - p))
End Function

Private Function SelectionIds(ByVal selection As String) As String()
    Dim body As String, p As Long, q As Long
    p = InStr(selection, ">"): q = InStr(selection, "</CLitSelection>")
    If p > 0 And q > p Then body = Mid$(selection, p + 1, q - p - 1)
    If Right$(body, 1) = "#" Then body = Left$(body, Len(body) - 1)
    SelectionIds = Split(body, "#")
End Function

Private Function NewIds(ByVal before As String, ByVal after As String) As Variant
    Dim known As Object, id As Variant, result() As String, n As Long
    Set known = CreateObject("Scripting.Dictionary")
    For Each id In SelectionIds(before): known(id) = True: Next id
    ReDim result(0 To 0)
    n = -1
    For Each id In SelectionIds(after)
        If Not known.Exists(id) Then
            n = n + 1
            ReDim Preserve result(0 To n)
            result(n) = id
        End If
    Next id
    If n < 0 Then NewIds = Array() Else NewIds = result
End Function

Private Function Sel(ByVal id As String) As String
    If id = "-1" Then
        Sel = "<CLitSelection Name="""" Ident=""-1"">-1</CLitSelection>"
    Else
        Sel = "<CLitSelection Name="""">" & id & "#</CLitSelection>"
    End If
End Function

Private Function Num(ByVal value As Double) As String
    Num = Trim$(Str$(Round(value, 6)))
End Function

Private Function XmlEscape(ByVal text As String) As String
    XmlEscape = Replace(Replace(Replace(Replace(text, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function

Private Function V(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Variant
    V = Array(x, y, z)
End Function

Private Function Add(a As Variant, b As Variant) As Variant
    Add = V(a(0) + b(0), a(1) + b(1), a(2) + b(2))
End Function

Private Function Subtract(a As Variant, b As Variant) As Variant
    Subtract = V(a(0) - b(0), a(1) - b(1), a(2) - b(2))
End Function

Private Function ScaleVec(a As Variant, ByVal k As Double) As Variant
    ScaleVec = V(a(0) * k, a(1) * k, a(2) * k)
End Function

Private Function Dot(a As Variant, b As Variant) As Double
    Dot = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function

Private Function Cross(a As Variant, b As Variant) As Variant
    Cross = V(a(1) * b(2) - a(2) * b(1), a(2) * b(0) - a(0) * b(2), a(0) * b(1) - a(1) * b(0))
End Function

Private Function Norm(a As Variant) As Double
    Norm = Sqr(Dot(a, a))
End Function

Private Function StripMetaPartNo(ByVal partNo As String) As String
    Dim startPos As Long, endPos As Long, result As String
    startPos = InStr(partNo, "Value=""")
    If startPos > 0 Then
        startPos = startPos + Len("Value=""")
        endPos = InStr(startPos, partNo, """")
        If endPos > startPos Then result = Mid(partNo, startPos, endPos - startPos)
    Else
        result = Replace(Replace(partNo, "<Meta.PartNo>", ""), "</Meta.PartNo>", "")
    End If
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, "")
    result = Replace(result, "&#13;", "")
    result = Replace(result, "&#10;", "")
    result = Replace(result, "&#9;", "")
    result = Replace(result, "&amp;", "&")
    StripMetaPartNo = Trim(result)
End Function

Private Sub LogPrint(ByVal logFile As Object, ByVal message As String)
    Debug.Print message
    If Not logFile Is Nothing Then logFile.WriteLine message
End Sub
