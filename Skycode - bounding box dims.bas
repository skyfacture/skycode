Option Compare Database
Option Explicit

' ============================================================
' Skycode - Best Fit Envelope Dimensions (SOLIDWORKS Composer)
' Adds length / width / height linear dimensions to every visible actor whose
' name contains "BoundingBox" or "BestFitEnvelope", case-insensitively (the proxy made by "Skycode - sldworks
' bounding box.bas").
'   - Box geometry comes from the proxy's Meta.BestFitOrigin and
'     Meta.BestFitEdge1..3 properties (corner and edge vectors, world mm).
'   - Edges are labelled by size: longest = length, middle = width,
'     shortest = height.
'   - Each dimension is offset by OFFSET_DISTANCE along the box axis that best
'     matches its world offset direction, so all leaders stay normal to their
'     edge and the length/width leaders share one corner.
'   - Of the parallel candidate edges, the one facing the camera is used.
'   - Dimensions are styled from the "Dims" style file and named
'     "SkycodeBBoxDim <role> <proxy id>"; rerunning replaces them.
'   - Matching actors are set to 100 alpha (opacity).
' Requires a VBA reference to composer.exe (DSComposer type library).
' ============================================================

Private Const LOG_FILE_PATH As String = "C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\boundingboxdims.txt"
Private Const STYLE_PATH As String = "C:\Users\skyfa\OneDrive\Documents\SOLIDWORKSComposer\User\Profiles\Dims.smgStyleSet"
Private Const PROXY_NAME_MATCH As String = "BoundingBox"
Private Const LEGACY_PROXY_NAME_MATCH As String = "BestFitEnvelope"
Private Const DIM_TAG As String = "SkycodeBBoxDim"
Private Const OFFSET_DISTANCE As Double = 100#
Private Const CMD_DIMENSIONS As Long = 178   ' crRCmdDimensions
Private Const VISIBILITY_VISIBLE As Long = 1

Public Sub BoundingBoxDimensions()
    Dim app As DSComposer.Application
    Dim doc As DSComposer.Document
    Dim fs As Object, logFile As Object
    Dim proxies As Variant, proxyId As Variant
    Dim oldSelection As String, styleProps As String, actorName As String
    Dim done As Long, skipped As Long, matched As Long

    LogOpen fs, logFile, LOG_FILE_PATH
    LogPrint logFile, Now & " - BoundingBoxDimensions starting..."
    On Error GoTo Failed

    Set app = New DSComposer.Application
    Set doc = app.GetActiveDocument
    If doc Is Nothing Then Err.Raise vbObjectError + 3001, , "No active Composer document."

    styleProps = LoadStyleProperties(STYLE_PATH)
    oldSelection = doc.Selection
    proxies = SelectionIds(doc.GetAllActors(0))   ' crAtvAllActors

    For Each proxyId In proxies
        actorName = doc.GetPropertyValue(Sel(CStr(proxyId)), "Actor.Name")
        If InStr(1, actorName, PROXY_NAME_MATCH, vbTextCompare) = 0 _
            And InStr(1, actorName, LEGACY_PROXY_NAME_MATCH, vbTextCompare) = 0 Then
            ' Neither proxy name occurs anywhere in this actor name.
        ElseIf InStr(1, actorName, DIM_TAG, vbBinaryCompare) > 0 Then
            ' Our own dimensions also contain the proxy name; ignore them.
        Else
            matched = matched + 1
            ApplyBoundingBoxOpacity doc, CStr(proxyId), logFile
            If doc.GetVisibility(Sel(CStr(proxyId))) <> VISIBILITY_VISIBLE Then
                LogPrint logFile, Now & " - Set opacity; skipped hidden actor " & proxyId
            ElseIf DimensionProxy(app, doc, CStr(proxyId), styleProps, logFile) Then
                done = done + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next proxyId

    doc.Selection = oldSelection
    If matched = 0 Then
        LogPrint logFile, Now & " - ERROR: BoundingBox or BestFitEnvelope actor not found. Run Skycode - sldworks bounding box.bas on the source model, then import or re-import the generated BoundingBox proxy into Composer."
        GoTo Cleanup
    End If
    LogPrint logFile, Now & " - Finished. Dimensioned: " & done & ", skipped: " & skipped
    GoTo Cleanup

Failed:
    LogPrint logFile, Now & " - ERROR " & Err.Number & ": " & Err.Description
    MsgBox "Best Fit envelope dimensions failed:" & vbCrLf & Err.Description, vbExclamation
Cleanup:
    If Not logFile Is Nothing Then logFile.Close
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
    Dim roundedLabel As String, labelXml As Object, labelNode As Object

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
    If AttrValue(doc.GetPropertyValue(Sel(id), "Actor.DisplayValue"), "Actor.DisplayValue", "Value") <> 0 Then
        Err.Raise vbObjectError + 3107, , "Composer did not disable the automatic dimension value: " & id
    End If
    Set labelXml = CreateObject("MSXML2.DOMDocument.6.0")
    labelXml.async = False
    If Not labelXml.LoadXML(doc.GetPropertyValue(Sel(id), "Text.MiddleString")) Then
        Err.Raise vbObjectError + 3107, , "Invalid dimension label XML: " & id & ": " & labelXml.parseError.reason
    End If
    Set labelNode = labelXml.SelectSingleNode("//Text.MiddleString/@Value")
    If labelNode Is Nothing Then
        Err.Raise vbObjectError + 3107, , "Composer returned no dimension label: " & id
    End If
    If labelNode.Text <> roundedLabel Then
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

' Style file properties, minus profile bookkeeping, in the explicit-close form SetPropertyMap accepts.
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

' First occurrence of attr on the first <tag ...> element in xml.
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

Private Function SelectionIds(ByVal selection As String) As Variant
    Dim body As String, p As Long, q As Long
    p = InStr(selection, ">"): q = InStr(selection, "</CLitSelection>")
    If p > 0 And q > p Then body = Mid$(selection, p + 1, q - p - 1)
    If Right$(body, 1) = "#" Then body = Left$(body, Len(body) - 1)
    If Len(body) = 0 Then SelectionIds = Array() Else SelectionIds = Split(body, "#")
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
    Sel = "<CLitSelection Name="""">" & id & "#</CLitSelection>"
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

Private Sub LogOpen(ByRef fs As Object, ByRef logFile As Object, ByVal outputPath As String)
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(fs.GetParentFolderName(outputPath)) Then Set logFile = fs.CreateTextFile(outputPath, True)
End Sub

Private Sub LogPrint(ByVal logFile As Object, ByVal message As String)
    Debug.Print message
    If Not logFile Is Nothing Then logFile.WriteLine message
End Sub
