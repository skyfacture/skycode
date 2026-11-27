Option Explicit

' Requires a reference to the installed SOLIDWORKS type library.
' Native custom-plane box is kept as the Skycode_BoundingBox feature. The proxy uses its
' conservative axis-aligned extents, so its edges are parallel to the model X, Y and Z axes.
Private Const DOCUMENT_PART As Long = 1
Private Const DOCUMENT_ASSEMBLY As Long = 2
Private Const FEATURE_BOUNDING_BOX As Long = 114
Private Const BOUNDING_BOX_CUSTOM_PLANE As Long = 2
Private Const ALIGNMENT_PLANE_NAME As String = "Front Plane"
Private Const BOX_FEATURE_NAME As String = "Skycode_BoundingBox"
Private Const LOG_FOLDER As String = "C:\skycode\logs"
Private gDiagnostics As String


Public Sub SldworksBoundingBox()
    Call GetCurrentModelBoundingBox
End Sub

Public Function GetCurrentModelBoundingBox( _
    Optional ByVal showMessage As Boolean = True) As String

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim boxFeature As SldWorks.Feature
    Dim boxDefinition As SldWorks.BoundingBoxFeatureData
    Dim boxSketch As SldWorks.Sketch
    Dim selectedPlane As Object
    Dim configuration As SldWorks.Configuration
    Dim modelType As Long
    Dim lengthMetres As Double
    Dim widthMetres As Double
    Dim heightMetres As Double
    Dim volumeMetres3 As Double
    Dim report As String
    Dim stage As String
    Dim createdFeature As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim sourcePath As String
    Dim proxyPath As String
    Dim reportPath As String
    Dim proxyComponent As SldWorks.Component2
    Dim fs As Object
    Dim textOutput As Object
    Dim selectAttempts As String

    On Error GoTo Failed

    stage = "Connecting to the running SOLIDWORKS session"
    Set swApp = AttachRunningSolidWorks()
    If swApp Is Nothing Then
        Err.Raise vbObjectError + 2150, , "SOLIDWORKS is not reachable from this macro host. " & _
            "Start SOLIDWORKS in the same Windows user session and elevation as this host, " & _
            "open the source model, and run the macro again. Nothing was changed."
    End If

    stage = "Getting the current open model"
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        Err.Raise vbObjectError + 2100, , "No model is open in SOLIDWORKS."
    End If

    modelType = swModel.GetType
    If modelType <> DOCUMENT_PART And modelType <> DOCUMENT_ASSEMBLY Then
        Err.Raise vbObjectError + 2101, , "The active document must be a part or assembly."
    End If
    Set configuration = swModel.ConfigurationManager.ActiveConfiguration
    sourcePath = swModel.GetPathName
    If Len(sourcePath) = 0 Then
        Err.Raise vbObjectError + 2116, , "Save the source model before generating proxy and report files."
    End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    proxyPath = fs.BuildPath(fs.GetParentFolderName(sourcePath), _
        fs.GetBaseName(sourcePath) & "_BoundingBox.SLDPRT")
    stage = "Choosing where to save the text report"
    reportPath = ChooseReportPath(fs.GetBaseName(proxyPath) & ".txt", fs.GetParentFolderName(sourcePath))
    If Len(reportPath) = 0 Then
        Debug.Print "Cancelled: no report location selected; no model changes or output files were made."
        Exit Function
    End If
    If StrComp(reportPath, sourcePath, vbTextCompare) = 0 _
        Or StrComp(reportPath, proxyPath, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 2140, , "The text report cannot replace the source model or proxy part."
    End If

    stage = "Checking for an existing root-model bounding box"
    Set boxFeature = FindBoundingBox(swModel)

    If Not boxFeature Is Nothing Then
        stage = "Removing the existing root-model bounding box"
        swModel.ClearSelection2 True
        If boxFeature.Select2(False, 0) = False Then
            Err.Raise vbObjectError + 2110, , "Could not select the existing bounding-box feature for overwrite."
        End If
        If swModel.Extension.DeleteSelection2(0) = False Then
            Err.Raise vbObjectError + 2111, , "Could not delete the existing bounding-box feature."
        End If
        swModel.ClearSelection2 True
        Set boxFeature = Nothing
    End If

    If boxFeature Is Nothing Then
        stage = "Defining the bounding box aligned to the Front, Top and Right planes"
        If Not SelectAlignmentPlane(swApp, swModel, ALIGNMENT_PLANE_NAME, selectAttempts) Then
            Err.Raise vbObjectError + 2106, , "Could not select '" & ALIGNMENT_PLANE_NAME & "'. Attempts: " & selectAttempts
        End If
        stage = "Creating the bounding-box definition"
        Set boxDefinition = swModel.FeatureManager.CreateDefinition(FEATURE_BOUNDING_BOX)
        If boxDefinition Is Nothing Then
            Err.Raise vbObjectError + 2102, , "SOLIDWORKS did not create a bounding-box definition."
        End If

        stage = "Setting the custom-plane reference type"
        boxDefinition.ReferenceFaceOrPlane = BOUNDING_BOX_CUSTOM_PLANE
        stage = "Accessing selections for the bounding-box definition"
        boxDefinition.AccessSelections swModel, Nothing
        stage = "Reading the selected alignment plane"
        Set selectedPlane = swModel.SelectionManager.GetSelectedObject6(1, -1)
        If selectedPlane Is Nothing Then
            Err.Raise vbObjectError + 2107, , "No plane is selected after AccessSelections (selection count " & _
                SelectedCount(swModel) & ")."
        End If
        stage = "Assigning the alignment plane as the planar entity"
        boxDefinition.PlanarEntity = selectedPlane
        stage = "Releasing selection access"
        boxDefinition.ReleaseSelectionAccess
        boxDefinition.IncludeHiddenBodies = True
        boxDefinition.IncludeSurfaces = False
        If modelType = DOCUMENT_ASSEMBLY Then
            boxDefinition.IncludeHiddenComponents = True
            boxDefinition.IncludeEnvelopeComponents = False
        End If

        stage = "Calculating the aligned bounding box"
        swModel.ClearSelection2 True
        Set boxFeature = swModel.FeatureManager.CreateFeature(boxDefinition)
        If boxFeature Is Nothing Then
            Err.Raise vbObjectError + 2103, , _
                "SOLIDWORKS did not create the bounding-box feature. Check that the model has usable solid geometry."
        End If
        createdFeature = True
        boxFeature.Name = BOX_FEATURE_NAME
        swModel.ClearSelection2 True
    End If

    stage = "Rebuilding the active model"
    If swModel.EditRebuild3 = False Then
        Err.Raise vbObjectError + 2107, , "The model rebuild failed; envelope dimensions have not been accepted."
    End If

    stage = "Reading the oriented bounding-box sketch"
    Set boxSketch = boxFeature.GetSpecificFeature2
    If boxSketch Is Nothing Then
        Err.Raise vbObjectError + 2108, , "The bounding-box sketch is unavailable in the active configuration."
    End If
    ' The native box can rotate within the Front Plane. Its axis-aligned sketch extents
    ' conservatively contain the source geometry without traversing every solid body.
    stage = "Computing the orthogonal (axis-aligned) envelope"
    Dim boxOrigin(0 To 2) As Double
    Dim boxLengths(0 To 2) As Double
    ComputeOrthogonalBox swApp, boxSketch, boxOrigin, boxLengths
    SortDimensions boxLengths, lengthMetres, widthMetres, heightMetres
    volumeMetres3 = lengthMetres * widthMetres * heightMetres

    stage = "Creating and saving the aligned solid proxy part"
    CreateEnvelopeProxy swApp, swModel, boxOrigin, boxLengths, proxyPath

    If modelType = DOCUMENT_ASSEMBLY Then
        stage = "Inserting the proxy as an envelope component"
        Dim assembly As SldWorks.AssemblyDoc
        Dim mathUtility As SldWorks.MathUtility
        Dim identity(0 To 15) As Double
        Dim identityData As Variant
        Set assembly = swModel
        Set proxyComponent = FindProxyComponent(assembly, proxyPath)
        If proxyComponent Is Nothing Then
            Set proxyComponent = assembly.InsertEnvelope(proxyPath, "", 0#, 0#, 0#)
        End If
        If proxyComponent Is Nothing Then
            Err.Raise vbObjectError + 2118, , "The proxy was saved, but envelope insertion failed."
        End If
        identity(0) = 1#: identity(4) = 1#: identity(8) = 1#: identity(12) = 1#
        identityData = identity
        Set mathUtility = swApp.GetMathUtility
        proxyComponent.Transform2 = mathUtility.CreateTransform(identityData)
        If proxyComponent.IsEnvelope = False Then
            Err.Raise vbObjectError + 2119, , "The inserted proxy is not marked as an envelope."
        End If
        proxyComponent.ExcludeFromBOM = True
        swModel.ClearSelection2 True
        If proxyComponent.Select4(False, Nothing, False) = False Then
            Err.Raise vbObjectError + 2120, , "Could not select the proxy component to fix its placement."
        End If
        assembly.FixComponent
        swModel.ClearSelection2 True
    End If
    If swModel.EditRebuild3 = False Then
        Err.Raise vbObjectError + 2122, , "The source model failed to rebuild after proxy creation."
    End If

    report = "Model: " & swModel.GetTitle & vbCrLf & _
        "Configuration: " & configuration.Name & vbCrLf & _
        "Feature: " & boxFeature.Name & vbCrLf & _
        "Conservative rectilinear envelope with sides orthogonal to the model X, Y and Z axes" & vbCrLf & _
        "Length: " & Format$(lengthMetres * 1000#, "0") & " mm" & vbCrLf & _
        "Width: " & Format$(widthMetres * 1000#, "0") & " mm" & vbCrLf & _
        "Height: " & Format$(heightMetres * 1000#, "0") & " mm" & vbCrLf & _
        "Box volume: " & Format$(volumeMetres3, "0.000000000") & " m^3" & vbCrLf & _
        "Box volume: " & Format$(volumeMetres3 * 1000#, "0.000000") & " litres" & vbCrLf & _
        "Dimensions are longest to shortest, not world X/Y/Z." & vbCrLf & _
        "Hidden solids included; surfaces and envelope components excluded." & vbCrLf & _
        "Proxy part: " & proxyPath & vbCrLf & _
        "Text report: " & reportPath & vbCrLf & _
        "The proxy part is saved; the source model has not been saved."

    If modelType = DOCUMENT_ASSEMBLY Then
        report = report & vbCrLf & _
            "Proxy present as a fixed envelope component, excluded from BOM." & vbCrLf & _
            "Enable envelope-component import in Composer."
    End If

    stage = "Writing the popup report to the selected file"
    Set textOutput = fs.CreateTextFile(reportPath, True, True)
    textOutput.Write report
    textOutput.Close
    Set textOutput = Nothing

    Debug.Print report
    GetCurrentModelBoundingBox = report
    If showMessage Then MsgBox report, vbInformation, "SOLIDWORKS Best Fit envelope"
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = stage & ": " & Err.Description
    If stage <> "Connecting to the running SOLIDWORKS session" Then
        failureDescription = failureDescription & _
            " Outputs or an inserted component may remain after a partial failure. Proxy: " & proxyPath
    End If
    If createdFeature Then
        failureDescription = failureDescription & _
            " A new bounding-box feature was added before the failure; inspect it in the unsaved model."
    End If
    Debug.Print "ERROR: " & failureDescription
    AppendFailureLog swModel, failureNumber, failureDescription
    Err.Raise failureNumber, "GetCurrentModelBoundingBox", failureDescription
End Function

Private Sub AppendFailureLog(ByVal model As SldWorks.ModelDoc2, ByVal number As Long, ByVal description As String)
    Dim fs As Object, stream As Object
    Dim title As String
    On Error Resume Next
    If Not model Is Nothing Then title = model.GetPathName
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(LOG_FOLDER) Then fs.CreateFolder LOG_FOLDER
    Set stream = fs.OpenTextFile(LOG_FOLDER & "\sldworks bounding box.txt", 8, True, -1)
    stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & title & " | error " & number & " | " & description
    If Len(gDiagnostics) > 0 Then stream.WriteLine "    " & gDiagnostics
    stream.Close
End Sub

Private Function AttachRunningSolidWorks() As Object
    ' The unversioned ProgID is registered to the newest installed SOLIDWORKS, which may not be
    ' the running one, so the version-specific ProgIDs are tried instead.
    Dim progIds As Variant, progId As Variant, found As Object
    progIds = Array("SldWorks.Application.34", "SldWorks.Application.33")
    For Each progId In progIds
        On Error Resume Next
        Set found = Nothing
        Set found = GetObject(, CStr(progId))
        On Error GoTo 0
        If Not found Is Nothing Then
            Set AttachRunningSolidWorks = found
            Exit Function
        End If
    Next progId
End Function

Private Function SelectedCount(ByVal model As SldWorks.ModelDoc2) As Long
    SelectedCount = model.SelectionManager.GetSelectedObjectCount2(-1)
End Function

' Tries every selection route for the plane and logs each result, so a failure shows exactly what was attempted.
Private Function SelectAlignmentPlane(ByVal swApp As SldWorks.SldWorks, ByVal model As SldWorks.ModelDoc2, _
    ByVal planeName As String, ByRef attempts As String) As Boolean
    Dim errors As Long
    Dim feature As SldWorks.Feature
    Dim typeName As Variant
    Dim appendFlag As Variant
    Dim ok As Boolean

    attempts = ""
    If swApp.ActivateDoc3(model.GetTitle, False, 0, errors) Is Nothing Then
        attempts = "ActivateDoc3 failed (" & errors & "); "
    End If
    For Each appendFlag In Array(False, True)
        For Each typeName In Array("PLANE", "DATUMPLANE")
            model.ClearSelection2 True
            ok = model.Extension.SelectByID2(planeName, CStr(typeName), 0#, 0#, 0#, CBool(appendFlag), 0, Nothing, 0)
            attempts = attempts & "SelectByID2(" & typeName & ",append=" & appendFlag & ")=" & ok & "/" & SelectedCount(model) & "; "
            If ok And SelectedCount(model) > 0 Then
                SelectAlignmentPlane = True
                Exit Function
            End If
        Next typeName
    Next appendFlag

    Set feature = model.FeatureByName(planeName)
    If feature Is Nothing Then
        attempts = attempts & "FeatureByName not found; "
        Exit Function
    End If
    For Each appendFlag In Array(False, True)
        model.ClearSelection2 True
        ok = feature.Select2(CBool(appendFlag), 0)
        attempts = attempts & "Feature.Select2(append=" & appendFlag & ")=" & ok & "/" & SelectedCount(model) & "; "
        If ok And SelectedCount(model) > 0 Then
            SelectAlignmentPlane = True
            Exit Function
        End If
    Next appendFlag
End Function

' Report goes beside the source model and overwrites any previous report, without prompting.
Private Function ChooseReportPath(ByVal suggestedName As String, ByVal initialFolder As String) As String
    ChooseReportPath = initialFolder & "\" & suggestedName
End Function

Private Sub CreateEnvelopeProxy(ByVal app As SldWorks.SldWorks, _
    ByVal source As SldWorks.ModelDoc2, ByRef origin() As Double, ByRef lengths() As Double, _
    ByVal outputPath As String)

    Dim utility As SldWorks.MathUtility
    Dim directions(0 To 2, 0 To 2) As Double
    Dim transformData(0 To 15) As Double
    Dim boxData(0 To 8) As Double
    Dim data As Variant
    Dim body As SldWorks.Body2
    Dim modeler As SldWorks.Modeler
    Dim partModel As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim feature As SldWorks.Feature
    Dim i As Long, j As Long
    Dim template As String
    Dim errors As Long, warnings As Long
    Dim activated As SldWorks.ModelDoc2
    Dim openDoc As SldWorks.ModelDoc2
    Dim fs As Object

    Set utility = app.GetMathUtility
    ' Edges run along the model X, Y and Z axes from the minimum corner.
    directions(0, 0) = 1#: directions(1, 1) = 1#: directions(2, 2) = 1#
    For i = 0 To 2
        If lengths(i) <= 0# Then Err.Raise vbObjectError + 2124, , "Degenerate box edge."
        For j = 0 To 2
            transformData(i * 3 + j) = directions(i, j)
        Next j
        transformData(9 + i) = origin(i)
    Next i
    transformData(12) = 1#

    ' Modeler box input is face centre, axis, width, length and height, in metres.
    boxData(0) = lengths(0) / 2#: boxData(1) = lengths(1) / 2#
    boxData(5) = 1#
    boxData(6) = lengths(0): boxData(7) = lengths(1): boxData(8) = lengths(2)
    data = boxData
    Set modeler = app.GetModeler
    Set body = modeler.CreateBodyFromBox3(data)
    If body Is Nothing Then Err.Raise vbObjectError + 2127, , "Could not create the solid box body."
    data = transformData
    If body.ApplyTransform(utility.CreateTransform(data)) = False Then
        Err.Raise vbObjectError + 2128, , "Could not align the box body with the source sketch."
    End If

    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(outputPath) Then
        Set openDoc = app.GetOpenDocumentByName(outputPath)
        If Not openDoc Is Nothing Then app.CloseDoc openDoc.GetTitle
        fs.DeleteFile outputPath, True
    End If
    template = app.GetUserPreferenceStringValue(8)
    If Len(template) = 0 Then Err.Raise vbObjectError + 2129, , "Configure a default SOLIDWORKS part template."
    Set partModel = app.NewDocument(template, 0, 0#, 0#)
    If partModel Is Nothing Then Err.Raise vbObjectError + 2130, , "Could not open the proxy part template."
    Set activated = app.ActivateDoc3(partModel.GetTitle, False, 0, errors)
    If activated Is Nothing Or errors <> 0 Then
        Err.Raise vbObjectError + 2136, , "Could not activate the proxy part."
    End If
    Set part = partModel
    Set feature = part.CreateFeatureFromBody3(body, False, 0)
    If feature Is Nothing Then Err.Raise vbObjectError + 2131, , "Could not add the solid body to the proxy part."
    feature.Name = BOX_FEATURE_NAME
    WriteBoxFrameProperties partModel, origin, directions, lengths
    Dim saved As Boolean
    saved = partModel.Extension.SaveAs(outputPath, 0, 1, Nothing, errors, warnings)
    If saved = False Or errors <> 0 Then
        Err.Raise vbObjectError + 2132, , "Proxy save failed. Error flags: " & errors & "; warning flags: " & warnings
    End If
    If warnings <> 0 Then Debug.Print "Proxy save warning flags: " & warnings
    Set activated = app.ActivateDoc3(source.GetTitle, False, 0, errors)
    If activated Is Nothing Or errors <> 0 Then
        Err.Raise vbObjectError + 2133, , "Could not reactivate the source model. Error flags: " & errors
    End If
End Sub

' Stores the box corner and three edge vectors (model coordinates, mm) as file
' custom properties. Composer imports them as Meta.BestFitOrigin/Edge1..3 for
' "Skycode - bestfit envelope dimensions.bas".
Private Sub WriteBoxFrameProperties(ByVal partModel As SldWorks.ModelDoc2, _
    ByVal origin As Variant, ByRef directions() As Double, ByRef lengths() As Double)

    Dim properties As SldWorks.CustomPropertyManager
    Dim i As Long

    Set properties = partModel.Extension.CustomPropertyManager("")
    If properties Is Nothing Then Err.Raise vbObjectError + 2141, , "Proxy custom properties are unavailable."
    SetTextProperty properties, "BestFitOrigin", _
        VectorText(origin(0) * 1000#, origin(1) * 1000#, origin(2) * 1000#)
    For i = 0 To 2
        SetTextProperty properties, "BestFitEdge" & (i + 1), VectorText( _
            directions(i, 0) * lengths(i) * 1000#, _
            directions(i, 1) * lengths(i) * 1000#, _
            directions(i, 2) * lengths(i) * 1000#)
    Next i
End Sub

Private Sub SetTextProperty(ByVal properties As SldWorks.CustomPropertyManager, _
    ByVal name As String, ByVal value As String)
    ' 30 = swCustomInfoText, 2 = swCustomPropertyReplaceValue; 0 = swCustomInfoAddResult_AddedOrChanged
    If properties.Add3(name, 30, value, 2) <> 0 Then
        Err.Raise vbObjectError + 2142, , "Could not write proxy custom property " & name & "."
    End If
End Sub

Private Function VectorText(ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
    VectorText = Trim$(Str$(Round(x, 6))) & ";" & Trim$(Str$(Round(y, 6))) & ";" & Trim$(Str$(Round(z, 6)))
End Function

Private Function FindProxyComponent(ByVal assembly As SldWorks.AssemblyDoc, _
    ByVal proxyPath As String) As SldWorks.Component2
    Dim components As Variant
    Dim component As SldWorks.Component2
    Dim i As Long
    components = assembly.GetComponents(True)
    If IsEmpty(components) Then Exit Function
    For i = LBound(components) To UBound(components)
        Set component = components(i)
        If StrComp(component.GetPathName, proxyPath, vbTextCompare) = 0 Then
            Set FindProxyComponent = component
            Exit Function
        End If
    Next i
End Function

Private Sub SortDimensions(ByRef lengths() As Double, ByRef longest As Double, _
    ByRef middle As Double, ByRef shortest As Double)
    Dim a As Double, b As Double, c As Double, t As Double
    a = lengths(0): b = lengths(1): c = lengths(2)
    If a < b Then t = a: a = b: b = t
    If b < c Then t = b: b = c: c = t
    If a < b Then t = a: a = b: b = t
    longest = a: middle = b: shortest = c
End Sub

' Uses the native bounding-box sketch's axis-aligned extents. This is conservative:
' the resulting box contains the native box and therefore all geometry it contains.
Private Sub ComputeOrthogonalBox(ByVal app As SldWorks.SldWorks, ByVal sketch As SldWorks.Sketch, _
    ByRef origin() As Double, ByRef lengths() As Double)

    Dim utility As SldWorks.MathUtility
    Dim limits(0 To 5) As Double
    Dim i As Long

    Set utility = app.GetMathUtility
    SketchExtents sketch, utility, limits
    gDiagnostics = "Envelope extents taken from the native bounding-box sketch; no body traversal performed."
    For i = 0 To 2
        origin(i) = limits(i)
        lengths(i) = limits(i + 3) - limits(i)
        If lengths(i) <= 0# Then
            Err.Raise vbObjectError + 2147, , "The native bounding-box sketch has a zero or negative extent."
        End If
    Next i
End Sub

Private Sub SketchExtents(ByVal sketch As SldWorks.Sketch, ByVal utility As SldWorks.MathUtility, _
    ByRef limits() As Double)

    Dim toModel As SldWorks.MathTransform
    Dim segments As Variant
    Dim line As SldWorks.SketchLine
    Dim point As Variant
    Dim i As Long, j As Long, k As Long
    Dim first As Boolean

    Set toModel = sketch.ModelToSketchTransform.Inverse
    segments = sketch.GetSketchSegments
    If Not IsArray(segments) Then
        Err.Raise vbObjectError + 2148, , "SOLIDWORKS did not return the bounding-box sketch edges."
    End If
    If UBound(segments) - LBound(segments) + 1 <> 12 Then
        Err.Raise vbObjectError + 2149, , "Expected exactly 12 edges in the bounding-box sketch."
    End If
    first = True
    For i = LBound(segments) To UBound(segments)
        Set line = segments(i)
        For k = 0 To 1
            If k = 0 Then
                point = SketchPointInModel(line.GetStartPoint2, utility, toModel)
            Else
                point = SketchPointInModel(line.GetEndPoint2, utility, toModel)
            End If
            For j = 0 To 2
                If first Or point(j) < limits(j) Then limits(j) = point(j)
                If first Or point(j) > limits(j + 3) Then limits(j + 3) = point(j)
            Next j
            first = False
        Next k
    Next i
End Sub

Private Function SketchPointInModel(ByVal point As SldWorks.SketchPoint, _
    ByVal utility As SldWorks.MathUtility, ByVal transform As SldWorks.MathTransform) As Variant
    Dim coordinates(0 To 2) As Double
    Dim data As Variant
    Dim mathPoint As SldWorks.MathPoint
    coordinates(0) = point.X: coordinates(1) = point.Y: coordinates(2) = point.Z
    data = coordinates
    Set mathPoint = utility.CreatePoint(data)
    Set mathPoint = mathPoint.MultiplyTransform(transform)
    SketchPointInModel = mathPoint.ArrayData
End Function

Private Function FindBoundingBox(ByVal model As SldWorks.ModelDoc2) As SldWorks.Feature
    Dim feature As SldWorks.Feature
    Dim found As SldWorks.Feature

    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        If feature.GetTypeName2 = "BoundingBoxProfileFeat" Then
            If Not found Is Nothing Then
                Err.Raise vbObjectError + 2109, , _
                    "More than one root-model bounding box was found; select the intended box before extending this script."
            End If
            Set found = feature
        End If
        Set feature = feature.GetNextFeature
    Loop
    Set FindBoundingBox = found
End Function
