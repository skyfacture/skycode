Option Explicit

' Requires a reference to the installed SOLIDWORKS type library.
' Native Best Fit is not a certified global-minimum-volume calculation.
Private Const DOCUMENT_PART As Long = 1
Private Const DOCUMENT_ASSEMBLY As Long = 2
Private Const FEATURE_BOUNDING_BOX As Long = 114
Private Const BOUNDING_BOX_BEST_FIT As Long = 1
Private Const SKETCH_SEGMENT_LINE As Long = 0
Private Const BOX_FEATURE_NAME As String = "Skycode_BoundingBox"

#If VBA7 Then
Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef dialog As SaveFileDialog) As Long
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
#Else
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef dialog As SaveFileDialog) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
#End If

Private Type SaveFileDialog
    structureSize As Long
#If VBA7 Then
    ownerWindow As LongPtr
    instance As LongPtr
    filter As LongPtr
    customFilter As LongPtr
#Else
    ownerWindow As Long
    instance As Long
    filter As Long
    customFilter As Long
#End If
    maxCustomFilter As Long
    filterIndex As Long
#If VBA7 Then
    file As LongPtr
#Else
    file As Long
#End If
    maxFile As Long
#If VBA7 Then
    fileTitle As LongPtr
#Else
    fileTitle As Long
#End If
    maxFileTitle As Long
#If VBA7 Then
    initialDirectory As LongPtr
    title As LongPtr
#Else
    initialDirectory As Long
    title As Long
#End If
    flags As Long
    fileOffset As Integer
    fileExtension As Integer
#If VBA7 Then
    defaultExtension As LongPtr
    customData As LongPtr
    hook As LongPtr
    templateName As LongPtr
    reserved As LongPtr
#Else
    defaultExtension As Long
    customData As Long
    hook As Long
    templateName As Long
    reserved As Long
#End If
    reservedValue As Long
    extendedFlags As Long
End Type

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

    On Error GoTo Failed

    stage = "Connecting to the running SOLIDWORKS session"
    Set swApp = GetObject(, "SldWorks.Application")

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

    If boxFeature Is Nothing Then
        stage = "Defining the native Best Fit envelope"
        Set boxDefinition = swModel.FeatureManager.CreateDefinition(FEATURE_BOUNDING_BOX)
        If boxDefinition Is Nothing Then
            Err.Raise vbObjectError + 2102, , "SOLIDWORKS did not create a bounding-box definition."
        End If

        boxDefinition.ReferenceFaceOrPlane = BOUNDING_BOX_BEST_FIT
        boxDefinition.IncludeHiddenBodies = True
        boxDefinition.IncludeSurfaces = False
        If modelType = DOCUMENT_ASSEMBLY Then
            boxDefinition.IncludeHiddenComponents = True
            boxDefinition.IncludeEnvelopeComponents = False
        End If

        stage = "Calculating the native Best Fit envelope"
        swModel.ClearSelection2 True
        Set boxFeature = swModel.FeatureManager.CreateFeature(boxDefinition)
        If boxFeature Is Nothing Then
            Err.Raise vbObjectError + 2103, , _
                "SOLIDWORKS did not create the bounding-box feature. Check that the model has usable solid geometry."
        End If
        createdFeature = True
        boxFeature.Name = BOX_FEATURE_NAME
        swModel.ClearSelection2 True
    Else
        stage = "Validating the existing bounding-box settings"
        Set boxDefinition = boxFeature.GetDefinition
        If boxDefinition Is Nothing Then
            Err.Raise vbObjectError + 2104, , "The existing bounding-box definition is unavailable."
        End If

        If boxDefinition.ReferenceFaceOrPlane <> BOUNDING_BOX_BEST_FIT _
            Or boxDefinition.IncludeHiddenBodies = False _
            Or boxDefinition.IncludeSurfaces <> False Then
            Err.Raise vbObjectError + 2105, , _
                "The existing bounding box uses different fit/body settings. It has not been overwritten."
        End If
        If modelType = DOCUMENT_ASSEMBLY Then
            If boxDefinition.IncludeHiddenComponents = False _
                Or boxDefinition.IncludeEnvelopeComponents <> False Then
                Err.Raise vbObjectError + 2106, , _
                    "The existing bounding box uses different component settings. It has not been overwritten."
            End If
        End If
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
    ReadBoxDimensions boxSketch, lengthMetres, widthMetres, heightMetres
    volumeMetres3 = lengthMetres * widthMetres * heightMetres

    stage = "Creating and saving the aligned solid proxy part"
    CreateEnvelopeProxy swApp, swModel, boxSketch, proxyPath

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
        "Native Best Fit envelope (not a certified global minimum)" & vbCrLf & _
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
    failureDescription = failureDescription & _
        " Outputs or an inserted component may remain after a partial failure. Proxy: " & proxyPath
    If createdFeature Then
        failureDescription = failureDescription & _
            " A new bounding-box feature was added before the failure; inspect it in the unsaved model."
    End If
    Debug.Print "ERROR: " & failureDescription
    Err.Raise failureNumber, "GetCurrentModelBoundingBox", failureDescription
End Function

Private Function ChooseReportPath(ByVal suggestedName As String, ByVal initialFolder As String) As String
    Dim dialog As SaveFileDialog
    Dim fileBuffer As String
    Dim filter As String
    Dim title As String
    Dim extension As String
    Dim dialogError As Long

    fileBuffer = suggestedName & String$(32768 - Len(suggestedName), vbNullChar)
    filter = "Text files (*.txt)" & vbNullChar & "*.txt" & vbNullChar & vbNullChar
    title = "Save Best Fit envelope text report"
    extension = "txt"
    dialog.structureSize = LenB(dialog)
    dialog.filter = StrPtr(filter)
    dialog.filterIndex = 1
    dialog.file = StrPtr(fileBuffer)
    dialog.maxFile = Len(fileBuffer)
    dialog.initialDirectory = StrPtr(initialFolder)
    dialog.title = StrPtr(title)
    dialog.defaultExtension = StrPtr(extension)
    ' Confirm overwrites, preserve the current directory, and require an existing folder.
    dialog.flags = &H2 Or &H8 Or &H800

    If GetSaveFileNameW(dialog) <> 0 Then
        ChooseReportPath = Left$(fileBuffer, InStr(fileBuffer, vbNullChar) - 1)
    Else
        dialogError = CommDlgExtendedError()
        If dialogError <> 0 Then
            Err.Raise vbObjectError + 2139, "ChooseReportPath", _
                "The report Save As dialog failed. Windows dialog error: " & dialogError
        End If
    End If
End Function

Private Sub CreateEnvelopeProxy(ByVal app As SldWorks.SldWorks, _
    ByVal source As SldWorks.ModelDoc2, ByVal sketch As SldWorks.Sketch, _
    ByVal outputPath As String)

    Dim utility As SldWorks.MathUtility
    Dim toModel As SldWorks.MathTransform
    Dim segments As Variant
    Dim line As SldWorks.SketchLine
    Dim origin As Variant
    Dim startPoint As Variant
    Dim endPoint As Variant
    Dim directions(0 To 2, 0 To 2) As Double
    Dim lengths(0 To 2) As Double
    Dim transformData(0 To 15) As Double
    Dim boxData(0 To 8) As Double
    Dim data As Variant
    Dim body As SldWorks.Body2
    Dim modeler As SldWorks.Modeler
    Dim partModel As SldWorks.ModelDoc2
    Dim part As SldWorks.PartDoc
    Dim feature As SldWorks.Feature
    Dim i As Long, j As Long, count As Long
    Dim determinant As Double
    Dim temporary As Double
    Dim template As String
    Dim errors As Long, warnings As Long
    Dim activated As SldWorks.ModelDoc2
    Dim oldFeature As SldWorks.Feature
    Dim fs As Object
    Dim existingPart As Boolean

    Set utility = app.GetMathUtility
    Set toModel = sketch.ModelToSketchTransform.Inverse
    segments = sketch.GetSketchSegments
    Set line = segments(LBound(segments))
    origin = SketchPointInModel(line.GetStartPoint2, utility, toModel)

    For i = LBound(segments) To UBound(segments)
        Set line = segments(i)
        startPoint = SketchPointInModel(line.GetStartPoint2, utility, toModel)
        endPoint = SketchPointInModel(line.GetEndPoint2, utility, toModel)
        If PointsCoincide(origin, endPoint) Then
            endPoint = startPoint
            startPoint = origin
        End If
        If PointsCoincide(origin, startPoint) Then
            If count >= 3 Then Err.Raise vbObjectError + 2123, , "Too many edges meet at the box corner."
            For j = 0 To 2
                directions(count, j) = endPoint(j) - origin(j)
                lengths(count) = lengths(count) + directions(count, j) ^ 2
            Next j
            lengths(count) = Sqr(lengths(count))
            If lengths(count) <= 0# Then Err.Raise vbObjectError + 2124, , "Degenerate box edge."
            For j = 0 To 2
                directions(count, j) = directions(count, j) / lengths(count)
            Next j
            count = count + 1
        End If
    Next i
    If count <> 3 Then Err.Raise vbObjectError + 2125, , "Could not identify three edges at the box corner."

    determinant = directions(0, 0) * (directions(1, 1) * directions(2, 2) - directions(1, 2) * directions(2, 1)) _
        - directions(0, 1) * (directions(1, 0) * directions(2, 2) - directions(1, 2) * directions(2, 0)) _
        + directions(0, 2) * (directions(1, 0) * directions(2, 1) - directions(1, 1) * directions(2, 0))
    If Abs(Abs(determinant) - 1#) > 0.0000001 Then
        Err.Raise vbObjectError + 2126, , "Box edge directions are not orthonormal."
    End If
    If determinant < 0# Then
        For j = 0 To 2
            temporary = directions(0, j)
            directions(0, j) = directions(1, j)
            directions(1, j) = temporary
        Next j
        temporary = lengths(0): lengths(0) = lengths(1): lengths(1) = temporary
    End If
    For i = 0 To 2
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
    existingPart = fs.FileExists(outputPath)
    If existingPart Then
        Set partModel = app.GetOpenDocumentByName(outputPath)
        If partModel Is Nothing Then
            Set partModel = app.OpenDoc6(outputPath, DOCUMENT_PART, 1, "", errors, warnings)
        End If
        If partModel Is Nothing Or errors <> 0 Then
            Err.Raise vbObjectError + 2134, , "Could not open the existing proxy. Error flags: " & errors
        End If
        Set oldFeature = partModel.FeatureByName(BOX_FEATURE_NAME)
        If oldFeature Is Nothing Then
            Err.Raise vbObjectError + 2135, , _
                "The existing output does not contain the generated proxy feature; it has not been overwritten."
        End If
    Else
        template = app.GetUserPreferenceStringValue(8)
        If Len(template) = 0 Then Err.Raise vbObjectError + 2129, , "Configure a default SOLIDWORKS part template."
        Set partModel = app.NewDocument(template, 0, 0#, 0#)
    End If
    If partModel Is Nothing Then Err.Raise vbObjectError + 2130, , "Could not open the proxy part template."
    Set activated = app.ActivateDoc3(partModel.GetTitle, False, 0, errors)
    If activated Is Nothing Or errors <> 0 Then
        Err.Raise vbObjectError + 2136, , "Could not activate the proxy part."
    End If
    If existingPart Then
        partModel.ClearSelection2 True
        If oldFeature.Select2(False, 0) = False Then
            Err.Raise vbObjectError + 2137, , "Could not select the previous proxy body."
        End If
        If partModel.Extension.DeleteSelection2(0) = False Then
            Err.Raise vbObjectError + 2138, , "Could not remove the previous proxy body."
        End If
    End If
    Set part = partModel
    Set feature = part.CreateFeatureFromBody3(body, False, 0)
    If feature Is Nothing Then Err.Raise vbObjectError + 2131, , "Could not add the solid body to the proxy part."
    feature.Name = BOX_FEATURE_NAME
    WriteBoxFrameProperties partModel, origin, directions, lengths
    Dim saved As Boolean
    If existingPart Then
        saved = partModel.Save3(1, errors, warnings)
    Else
        saved = partModel.Extension.SaveAs(outputPath, 0, 1, Nothing, errors, warnings)
    End If
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

Private Function PointsCoincide(ByVal first As Variant, ByVal second As Variant) As Boolean
    PointsCoincide = Abs(first(0) - second(0)) < 0.000000001 _
        And Abs(first(1) - second(1)) < 0.000000001 _
        And Abs(first(2) - second(2)) < 0.000000001
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

Private Sub ReadBoxDimensions(ByVal sketch As SldWorks.Sketch, _
    ByRef lengthMetres As Double, ByRef widthMetres As Double, _
    ByRef heightMetres As Double)

    Dim segments As Variant
    Dim segment As SldWorks.SketchSegment
    Dim edgeLengths(0 To 11) As Double
    Dim i As Long
    Dim j As Long
    Dim offset As Long
    Dim swapLength As Double
    Dim tolerance As Double

    segments = sketch.GetSketchSegments
    If IsEmpty(segments) Then
        Err.Raise vbObjectError + 2110, , "The bounding-box sketch has no edges."
    End If
    If Not IsArray(segments) Then
        Err.Raise vbObjectError + 2111, , "SOLIDWORKS did not return an array of bounding-box edges."
    End If
    If UBound(segments) - LBound(segments) + 1 <> 12 Then
        Err.Raise vbObjectError + 2112, , "Expected exactly 12 edges in the bounding-box sketch."
    End If

    For i = LBound(segments) To UBound(segments)
        Set segment = segments(i)
        If segment.GetType <> SKETCH_SEGMENT_LINE Then
            Err.Raise vbObjectError + 2113, , "The bounding-box sketch contains a non-linear edge."
        End If
        offset = i - LBound(segments)
        edgeLengths(offset) = segment.GetLength
        If edgeLengths(offset) <= 0# Then
            Err.Raise vbObjectError + 2114, , "The bounding box has a zero or negative edge length."
        End If
    Next i

    For i = 0 To 10
        For j = i + 1 To 11
            If edgeLengths(j) < edgeLengths(i) Then
                swapLength = edgeLengths(i)
                edgeLengths(i) = edgeLengths(j)
                edgeLengths(j) = swapLength
            End If
        Next j
    Next i

    ' Each box dimension occurs on four edges, including for cubes and square prisms.
    For i = 0 To 8 Step 4
        tolerance = 0.0000000001 + edgeLengths(i + 3) * 0.00000001
        If edgeLengths(i + 3) - edgeLengths(i) > tolerance Then
            Err.Raise vbObjectError + 2115, , "The sketch edge lengths do not form three groups of four."
        End If
    Next i

    heightMetres = edgeLengths(3)
    widthMetres = edgeLengths(7)
    lengthMetres = edgeLengths(11)
End Sub