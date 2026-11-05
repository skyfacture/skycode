Option Explicit

' Requires the installed SOLIDWORKS type library; no constants-library reference.
Private Const ASSEMBLY_DOCUMENT As Long = 2
Private Const DEFAULT_ASSEMBLY_TEMPLATE As Long = 9
Private Const MATRIX_TOLERANCE As Double = 0.00000001
Private Const PRODUCTION_CONFIGURATION As String = "Default"
Private Const KIT_CONFIGURATION As String = "kitified"

Private mApp As SldWorks.SldWorks
Private mFS As Object
Private mSource As Collection
Private mReference As Collection
Private mGroups As Object
Private mFolder As String
Private mRunId As String
Private mStage As String
Private mLog As Object
Private mWorking As SldWorks.ModelDoc2
Private mSourcePath As String
Private mChildPath As String
Private mPeakGeneratedDocuments As Long
Private mRootMetadata As Collection
Private mProductionRoots As Object
Private mKitRoots As Object

Public Sub SkycodeKitify()
    Dim result As String
    result = KitifyDrifter()
End Sub

Public Function KitifyDrifter( _
    Optional ByVal sourcePath As String = "", _
    Optional ByVal referencePath As String = "", _
    Optional ByVal outputFolder As String = "", _
    Optional ByVal showMessage As Boolean = True) As String

    Dim source As SldWorks.ModelDoc2
    Dim reference As SldWorks.ModelDoc2
    Dim original As SldWorks.ModelDoc2
    Dim result As String
    Dim snapshotPath As String
    Dim template As String
    Dim n As Object
    Dim key As Variant
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim errors As Long
    Dim warnings As Long
    Dim reopened As SldWorks.ModelDoc2
    On Error GoTo Failed

    mStage = "Connecting to the running SOLIDWORKS session"
    Initialise
    Set original = mApp.ActiveDoc
    mStage = "Finding the saved production drifter"
    Set source = FindSource(sourcePath)
    Set mRootMetadata = ReadMetadata(source, source.ConfigurationManager.ActiveConfiguration.Name)
    mSourcePath = source.GetPathName
    CheckVisibleDocuments
    If Len(referencePath) = 0 Then Set mReference = LoadKitifyRules(source)
    mFolder = NewOutputFolder(source.GetPathName, outputFolder)
    mRunId = Format$(Now, "yyyymmdd_hhnnss") & "_" & Hex$(CLng(Timer * 1000#))
    Set mLog = mFS.CreateTextFile(mFS.BuildPath(mFolder, "run.txt"), False, True)
    LogLine "Source: " & source.GetPathName
    LogLine "Source configuration: " & source.ConfigurationManager.ActiveConfiguration.Name
    LogLine "Source initially dirty: " & CStr(source.GetSaveFlag)

    mStage = "Mapping every production occurrence before planning kits"
    Set mSource = Snapshot(source)
    WriteTree mSource, "source-tree.csv"
    WriteLocations

    mStage = "Reading saved reference rules"
    If Len(referencePath) > 0 Then
        Set reference = FindReference(source, referencePath)
        If StrComp(source.GetPathName, reference.GetPathName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 2400, , "Source and reference must be different documents."
        End If
        If reference.Visible Then Err.Raise vbObjectError + 2440, , _
            "Close the development reference and use saved rules to keep only two visible documents."
        LogLine "Reference: " & reference.GetPathName
        LogLine "Reference configuration: " & reference.ConfigurationManager.ActiveConfiguration.Name
        LogLine "Reference initially dirty: " & CStr(reference.GetSaveFlag)
        Set mReference = Snapshot(reference)
    Else
        LogLine "Saved reference membership rules: " & KitifyRulesPath(source.GetPathName)
        LogLine "Reference geometry was not opened. Rule-tree transforms are placeholders, not reference geometry."
    End If
    WriteTree mReference, "reference-tree.csv"

    mStage = "Learning kit hierarchy and matching source occurrences"
    Set mGroups = Dictionary()
    AddGroup "ROOT", "", "ROOT", "", ""
    Set mGroups("ROOT")("metadata") = mRootMetadata
    mGroups("ROOT")("sourcePath") = source.GetPathName
    mGroups("ROOT")("sourceConfig") = source.ConfigurationManager.ActiveConfiguration.Name
    DiscoverKits
    MatchOccurrences
    AssignDestinations
    CollectKits
    PrepareKitMetadata
    WritePlan

    template = mApp.GetUserPreferenceStringValue(DEFAULT_ASSEMBLY_TEMPLATE)
    If mFS.FileExists(template) = False Then
        Err.Raise vbObjectError + 2401, , "Set a valid default assembly template in SOLIDWORKS."
    End If
    Set mWorking = mApp.NewDocument(template, 0, 0#, 0#)
    If mWorking Is Nothing Then Err.Raise vbObjectError + 2416, , "Could not create the working kitified assembly."
    mPeakGeneratedDocuments = 1
    CheckVisibleDocuments
    mStage = "Building independent assembly documents inside-out"
    result = BuildGroup("ROOT", template, mFS.GetBaseName(source.GetPathName))

    mStage = "Closing and reopening saved output for persistence verification"
    CloseGenerated
    Set reopened = mApp.OpenDoc6(result, ASSEMBLY_DOCUMENT, 1, "", errors, warnings)
    If reopened Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2434, , _
        "Could not reopen saved output; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "REOPEN WARNING: " & warnings
    Set mWorking = reopened
    CheckVisibleDocuments
    mStage = "Verifying reopened output against the source and reference"
    ValidateOutput result, "snapshot-"
    ValidateMetadata "snapshot-"
    ValidateAppearances
    mStage = "Cloning production and combining configuration-specific trees"
    snapshotPath = result
    result = CombineConfigurations(source)
    mStage = "Reopening and verifying both saved configurations"
    CloseGenerated
    errors = 0
    warnings = 0
    Set reopened = mApp.OpenDoc6(result, ASSEMBLY_DOCUMENT, 1, KIT_CONFIGURATION, errors, warnings)
    If reopened Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Could not reopen combined output; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "COMBINED REOPEN WARNING: " & warnings
    Set mWorking = reopened
    CheckVisibleDocuments
    ValidateConfigurations
    For Each n In mSource
        n("validated") = False
    Next n
    ValidateOutput result
    ValidateMetadata
    If Not mApp.GetOpenDocumentByName(snapshotPath) Is Nothing Then _
        Err.Raise vbObjectError + 2453, , "Construction snapshot is unexpectedly still loaded: " & snapshotPath
    mFS.DeleteFile snapshotPath, True
    LogLine "Removed verified intermediate construction snapshot: " & snapshotPath
    For Each key In mGroups.Keys
        Set n = mGroups(key)
        If key <> "ROOT" And Len(n("file")) > 0 Then mApp.CloseDoc mFS.GetFileName(n("file"))
    Next key
    LogLine "Source finally dirty: " & CStr(source.GetSaveFlag)
    If Not reference Is Nothing Then LogLine "Reference finally dirty: " & CStr(reference.GetSaveFlag)
    LogLine "Document budget PASS: maximum 2 visible documents; peak separately opened generated documents=" & _
        CStr(mPeakGeneratedDocuments) & "; saved children closed after each insertion."
    LogLine "Referenced child geometry remains loaded internally by SOLIDWORKS."
    LogLine "VERIFIED: " & result
    LogLine "Source/reference files were not saved or restructured. See comparison.csv for reference differences."
    mLog.Close
    Set mLog = Nothing
    If Not original Is Nothing Then Activate original
    KitifyDrifter = result
    If showMessage Then MsgBox "Kitified assembly verified:" & vbCrLf & result & vbCrLf & _
        "Reports: " & mFolder & vbCrLf & _
        "Reference differences are recorded in comparison.csv.", vbInformation, "Skycode kitify"
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = mStage & ": " & Err.Description & vbCrLf & _
        "Partial outputs may remain in: " & mFolder & vbCrLf & _
        "Source/reference files have not been saved by this macro."
    Debug.Print failureDescription
    If Len(mChildPath) > 0 Then
        mApp.CloseDoc mChildPath
        mChildPath = ""
    End If
    If Not mLog Is Nothing Then
        mLog.WriteLine "FAILED: " & failureDescription
        mLog.Close
        Set mLog = Nothing
    End If
    Err.Raise failureNumber, "Skycode.KitifyDrifter", failureDescription
End Function

Private Sub Initialise()
    Set mApp = GetObject(, "SldWorks.Application")
    Set mFS = CreateObject("Scripting.FileSystemObject")
    Set mSource = Nothing
    Set mReference = Nothing
    Set mGroups = Nothing
    Set mLog = Nothing
    Set mWorking = Nothing
    mSourcePath = ""
    mChildPath = ""
    mPeakGeneratedDocuments = 0
    Set mRootMetadata = Nothing
    Set mProductionRoots = Nothing
    Set mKitRoots = Nothing
    mFolder = ""
    mRunId = ""
End Sub

Private Function Dictionary() As Object
    Set Dictionary = CreateObject("Scripting.Dictionary")
    Dictionary.CompareMode = vbTextCompare
End Function

Private Function FindSource(ByVal path As String) As SldWorks.ModelDoc2
    Dim doc As SldWorks.ModelDoc2
    Dim found As SldWorks.ModelDoc2
    Dim pattern As Object
    Dim assemblies As New Collection
    Dim names As Object
    Dim item As Variant
    Set pattern = CreateObject("VBScript.RegExp")
    Set names = Dictionary()
    pattern.Pattern = "^[0-9]{3,4} [0-9]{3,4} [0-9]{2}$"
    Set doc = mApp.GetFirstDocument
    Do While Not doc Is Nothing
        If doc.GetType = ASSEMBLY_DOCUMENT And Len(doc.GetPathName) > 0 Then
            assemblies.Add doc
            names(mFS.GetBaseName(doc.GetPathName)) = True
        End If
        Set doc = doc.GetNext
    Loop
    For Each item In assemblies
        Set doc = item
        If (Len(path) > 0 And StrComp(doc.GetPathName, path, vbTextCompare) = 0) Or _
            (Len(path) = 0 And pattern.Test(mFS.GetBaseName(doc.GetPathName)) And _
                (names.Exists(mFS.GetBaseName(doc.GetPathName) & " kitified") Or _
                    mFS.FileExists(KitifyRulesPath(doc.GetPathName)))) Then
            If Not found Is Nothing Then Err.Raise vbObjectError + 2402, , _
                "Multiple production drifters are open; pass the intended sourcePath."
            Set found = doc
        End If
    Next item
    If found Is Nothing Then Err.Raise vbObjectError + 2403, , _
        "Open the saved [part number].sldasm production drifter first."
    Set FindSource = found
End Function

Private Function KitifyRulesPath(ByVal sourcePath As String) As String
    KitifyRulesPath = mFS.BuildPath(mFS.GetParentFolderName(sourcePath), _
        mFS.GetBaseName(sourcePath) & " kitify-rules.xml")
End Function

Private Function KitifyRuleFields() As Variant
    KitifyRuleFields = Array("id", "parent", "instance", "path", "config", "note", _
        "assembly", "suppressed", "envelope", "excluded")
End Function

Public Sub LearnKitifyRules()
    Dim source As SldWorks.ModelDoc2
    Dim reference As SldWorks.ModelDoc2
    Dim xml As Object
    Dim root As Object
    Dim element As Object
    Dim field As Variant
    Dim n As Object
    Dim path As String
    Initialise
    Set source = FindSource("")
    Set reference = FindReference(source, "")
    Set mReference = Snapshot(reference)
    Set mGroups = Dictionary()
    AddGroup "ROOT", "", "ROOT", "", ""
    DiscoverKits
    path = KitifyRulesPath(source.GetPathName)
    If mFS.FileExists(path) Then Err.Raise vbObjectError + 2441, , _
        "Rules already exist; move them aside explicitly before relearning: " & path
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set root = xml.createElement("KitifyRules")
    root.setAttribute "version", "1"
    root.setAttribute "source", Identity(source.GetPathName)
    root.setAttribute "configuration", source.ConfigurationManager.ActiveConfiguration.Name
    xml.appendChild root
    For Each n In mReference
        Set element = xml.createElement("Occurrence")
        For Each field In KitifyRuleFields()
            element.setAttribute CStr(field), CStr(n(field))
        Next field
        root.appendChild element
    Next n
    xml.Save path
    Activate source
    MsgBox "Saved kit rules. Close the development reference before normal runs:" & vbCrLf & path, _
        vbInformation, "Skycode kitify"
End Sub

Private Function LoadKitifyRules(ByVal source As SldWorks.ModelDoc2) As Collection
    Dim xml As Object
    Dim root As Object
    Dim element As Object
    Dim field As Variant
    Dim n As Object
    Dim nodes As New Collection
    Dim path As String
    path = KitifyRulesPath(source.GetPathName)
    If Not mFS.FileExists(path) Then Err.Raise vbObjectError + 2442, , _
        "Missing saved kit rules: " & path & vbCrLf & _
        "Run LearnKitifyRules once with the development reference open, then close it."
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    xml.resolveExternals = False
    xml.setProperty "ProhibitDTD", True
    If xml.Load(path) = False Then Err.Raise vbObjectError + 2443, , _
        "Cannot load kit rules: " & xml.parseError.reason
    Set root = xml.documentElement
    If root Is Nothing Then Err.Raise vbObjectError + 2443, , "Empty kit rules: " & path
    If root.nodeName <> "KitifyRules" Or root.getAttribute("version") <> "1" Or _
        root.getAttribute("source") <> Identity(source.GetPathName) Or _
        root.getAttribute("configuration") <> source.ConfigurationManager.ActiveConfiguration.Name Then _
        Err.Raise vbObjectError + 2443, , "Kit rules do not match this source/configuration: " & path
    For Each element In root.childNodes
        If element.nodeType = 1 Then
            If element.nodeName <> "Occurrence" Then Err.Raise vbObjectError + 2443, , "Unexpected kit rule element."
            Set n = Dictionary()
            For Each field In KitifyRuleFields()
                If element.getAttributeNode(CStr(field)) Is Nothing Then Err.Raise vbObjectError + 2443, , _
                    "Missing kit rule field: " & field
                n.Add CStr(field), CStr(element.getAttribute(CStr(field)))
            Next field
            n("id") = CLng(n("id"))
            n("parent") = CLng(n("parent"))
            If n("id") <> nodes.Count + 1 Or n("parent") < 0 Or n("parent") >= n("id") Then _
                Err.Raise vbObjectError + 2443, , "Invalid kit rule ancestry."
            For Each field In Array("assembly", "suppressed", "envelope", "excluded")
                n(field) = CBool(n(field))
            Next field
            n.Add "identity", Identity(CStr(n("path")))
            n.Add "partKey", PartKey(CStr(n("identity")))
            n.Add "category", Category(CStr(n("note")))
            n.Add "match", 0&
            n.Add "kit", ""
            n.Add "fixed", False
            n.Add "matrix", IdentityMatrix()
            nodes.Add n
        End If
    Next element
    If nodes.Count = 0 Then Err.Raise vbObjectError + 2443, , "No saved reference occurrences."
    Set LoadKitifyRules = nodes
End Function

Private Function FindReference(ByVal source As SldWorks.ModelDoc2, ByVal path As String) As SldWorks.ModelDoc2
    Dim doc As SldWorks.ModelDoc2
    Dim found As SldWorks.ModelDoc2
    Dim expected As String
    expected = mFS.GetBaseName(source.GetPathName) & " kitified"
    Set doc = mApp.GetFirstDocument
    Do While Not doc Is Nothing
        If doc.GetType = ASSEMBLY_DOCUMENT And Len(doc.GetPathName) > 0 Then
            If (Len(path) > 0 And StrComp(doc.GetPathName, path, vbTextCompare) = 0) Or _
                (Len(path) = 0 And StrComp(mFS.GetBaseName(doc.GetPathName), expected, vbTextCompare) = 0) Then
                If Not found Is Nothing Then Err.Raise vbObjectError + 2404, , "Reference selection is ambiguous."
                Set found = doc
            End If
        End If
        Set doc = doc.GetNext
    Loop
    If found Is Nothing Then Err.Raise vbObjectError + 2405, , _
        "Open " & expected & ".sldasm first, or pass referencePath."
    Set FindReference = found
End Function

Private Function NewOutputFolder(ByVal sourcePath As String, ByVal requested As String) As String
    Dim parent As String
    Dim candidate As String
    Dim suffix As Long
    Dim folderName As String
    If Len(requested) > 0 Then
        candidate = mFS.GetAbsolutePathName(requested)
        If mFS.FolderExists(candidate) Or mFS.FileExists(candidate) Then
            Err.Raise vbObjectError + 2406, , "Output folder must not already exist: " & candidate
        End If
    Else
        parent = mFS.GetParentFolderName(sourcePath)
        folderName = "kitify-generated_" & Format$(Now, "yyyymmdd_hhnnss")
        candidate = mFS.BuildPath(parent, folderName)
        Do While mFS.FolderExists(candidate) Or mFS.FileExists(candidate)
            suffix = suffix + 1
            candidate = mFS.BuildPath(parent, folderName & "_" & CStr(suffix))
        Loop
    End If
    mFS.CreateFolder candidate
    NewOutputFolder = candidate
End Function

Private Function Snapshot(ByVal model As SldWorks.ModelDoc2) As Collection
    Dim assembly As SldWorks.AssemblyDoc
    Dim root As SldWorks.Component2
    Dim nodes As New Collection
    Dim status As Long
    Set assembly = model
    Activate model
    status = assembly.ResolveAllLightWeightComponents(False)
    LogLine "Resolve lightweight status (" & model.GetTitle & "): " & CStr(status)
    Set root = model.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    If root Is Nothing Then Err.Raise vbObjectError + 2407, , "Assembly root is unavailable: " & model.GetTitle
    WalkChildren root, 0, False, nodes
    If nodes.Count = 0 Then Err.Raise vbObjectError + 2408, , "Assembly tree is empty: " & model.GetTitle
    Set Snapshot = nodes
End Function

Private Sub WalkChildren(ByVal parent As SldWorks.Component2, ByVal parentId As Long, _
    ByVal excluded As Boolean, ByVal nodes As Collection)
    Dim children As Variant
    Dim item As Variant
    Dim component As SldWorks.Component2
    Dim model As SldWorks.ModelDoc2
    Dim transform As SldWorks.MathTransform
    Dim n As Object
    Dim id As Long
    children = parent.GetChildren
    If IsEmpty(children) Then Exit Sub
    For Each item In children
        Set component = item
        Set n = Dictionary()
        id = nodes.Count + 1
        n.Add "id", id
        n.Add "parent", parentId
        n.Add "instance", component.Name2
        n.Add "path", component.GetPathName
        n.Add "identity", Identity(CStr(n("path")))
        n.Add "partKey", PartKey(CStr(n("identity")))
        n.Add "config", component.ReferencedConfiguration
        n.Add "suppressed", CBool(component.IsSuppressed <> False)
        n.Add "envelope", CBool(component.IsEnvelope <> False)
        n.Add "excluded", CBool(excluded Or n("suppressed") Or n("envelope"))
        n.Add "assembly", CBool(LCase$(mFS.GetExtensionName(n("path"))) = "sldasm")
        n.Add "note", ""
        n.Add "category", ""
        n.Add "match", 0&
        n.Add "kit", ""
        n.Add "destination", ""
        n.Add "validated", False
        n.Add "fixed", CBool(component.IsFixed <> False)
        n.Add "reason", ""
        Set model = component.GetModelDoc2
        If n("excluded") = False Then
            If model Is Nothing Then Err.Raise vbObjectError + 2409, , _
                "Unresolved active component: " & component.Name2
            If model.GetType <> 1 And model.GetType <> ASSEMBLY_DOCUMENT Then _
                Err.Raise vbObjectError + 2410, , "Unsupported component: " & component.Name2
            If Len(n("path")) = 0 Or mFS.FileExists(n("path")) = False Then _
                Err.Raise vbObjectError + 2411, , "Save/resolve external component: " & component.Name2
        End If
        If Not model Is Nothing Then n("note") = ReadNote(model, CStr(n("config")))
        If n("assembly") And n("excluded") = False Then _
            Set n("metadata") = ReadMetadata(model, CStr(n("config")))
        n("category") = Category(CStr(n("note")))
        Set transform = component.Transform2
        If transform Is Nothing Then
            If n("excluded") = False Then Err.Raise vbObjectError + 2412, , _
                "No installed transform: " & component.Name2
            n.Add "matrix", IdentityMatrix()
        Else
            n.Add "matrix", transform.ArrayData
        End If
        nodes.Add n
        WalkChildren component, id, CBool(n("excluded")), nodes
    Next item
End Sub

Private Function ReadNote(ByVal model As SldWorks.ModelDoc2, ByVal config As String) As String
    Dim manager As SldWorks.CustomPropertyManager
    Dim raw As String
    Dim resolved As String
    Dim wasResolved As Boolean
    Dim linked As Boolean
    Set manager = model.Extension.CustomPropertyManager(config)
    manager.Get6 "NOTE", False, raw, resolved, wasResolved, linked
    If Len(Trim$(resolved)) = 0 Then resolved = raw
    If Len(Trim$(resolved)) = 0 And Len(config) > 0 Then
        Set manager = model.Extension.CustomPropertyManager("")
        manager.Get6 "NOTE", False, raw, resolved, wasResolved, linked
        If Len(Trim$(resolved)) = 0 Then resolved = raw
    End If
    ReadNote = Trim$(resolved)
End Function

Private Function IsSourceProperty(ByVal name As String) As Boolean
    Select Case UCase$(name)
        Case "SKYCODE.KIND", "SKYCODE.CATEGORY", "SKYCODE.PROVENANCE", "SKYCODE.GENERATED"
            IsSourceProperty = False
        Case Else
            IsSourceProperty = True
    End Select
End Function

Private Function ReadMetadata(ByVal model As SldWorks.ModelDoc2, ByVal config As String) As Collection
    Dim result As New Collection
    Dim manager As SldWorks.CustomPropertyManager
    Dim scope As Variant
    Dim names As Variant
    Dim name As Variant
    Dim entry As Object
    Dim raw As String
    Dim resolved As String
    Dim wasResolved As Boolean
    Dim linked As Boolean
    Dim status As Long
    For Each scope In Array("", config)
        Set manager = model.Extension.CustomPropertyManager(CStr(scope))
        names = manager.GetNames
        If IsArray(names) Then
            For Each name In names
                If IsSourceProperty(CStr(name)) Then
                    status = manager.Get6(CStr(name), False, raw, resolved, wasResolved, linked)
                    If status = 1 Or (linked And wasResolved = False) Then _
                        Err.Raise vbObjectError + 2448, , "Cannot resolve source metadata: " & _
                        model.GetPathName & "; configuration=" & config & "; property=" & name
                    If Len(resolved) = 0 Then resolved = raw
                    Set entry = Dictionary()
                    entry.Add "scope", CStr(IIf(Len(scope) = 0, "", "CONFIG"))
                    entry.Add "name", CStr(name)
                    entry.Add "type", manager.GetType2(CStr(name))
                    entry.Add "value", resolved
                    result.Add entry
                End If
            Next name
        End If
    Next scope
    Set ReadMetadata = result
End Function

Private Sub ApplyMetadata(ByVal model As SldWorks.ModelDoc2, ByVal metadata As Collection)
    Dim manager As SldWorks.CustomPropertyManager
    Dim scope As Variant
    Dim names As Variant
    Dim name As Variant
    Dim entry As Object
    Dim config As String
    config = model.ConfigurationManager.ActiveConfiguration.Name
    For Each scope In Array("", config)
        Set manager = model.Extension.CustomPropertyManager(CStr(scope))
        names = manager.GetNames
        If IsArray(names) Then
            For Each name In names
                If IsSourceProperty(CStr(name)) Then
                    If manager.Delete2(CStr(name)) <> 0 Then Err.Raise vbObjectError + 2449, , _
                        "Could not clear previous working-document metadata: " & name
                End If
            Next name
        End If
    Next scope
    For Each entry In metadata
        scope = ""
        If entry("scope") = "CONFIG" Then scope = config
        Set manager = model.Extension.CustomPropertyManager(CStr(scope))
        If manager.Add3(CStr(entry("name")), CLng(entry("type")), CStr(entry("value")), 2) <> 0 Then _
            Err.Raise vbObjectError + 2449, , "Could not preserve source metadata: " & entry("name")
    Next entry
End Sub

Private Sub ValidateMetadata(Optional ByVal reportPrefix As String = "", Optional ByVal rootOnly As Boolean = False)
    Dim report As Object
    Dim key As Variant
    Dim group As Object
    Dim model As SldWorks.ModelDoc2
    Dim actual As Collection
    Dim expected As Collection
    Dim values As Object
    Dim entry As Object
    Dim found As Object
    Dim propertyKey As String
    Dim name As Variant
    Dim present As Boolean
    Dim count As Long
    Dim missingStatus As String
    Set report = NewCsvReport()
    WriteCsvLine report, "Group,GeneratedAssembly,Scope,Property,SourceFile,SourceConfiguration,ExpectedValue,ActualValue,Type,Status"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("count") > 0 And (rootOnly = False Or key = "ROOT") Then
            Set model = mApp.GetOpenDocumentByName(group("file"))
            If model Is Nothing Then Err.Raise vbObjectError + 2450, , "Generated metadata document is not loaded: " & key
            Set expected = group("metadata")
            Set actual = ReadMetadata(model, CStr(group("configuration")))
            If actual.Count <> expected.Count Then Err.Raise vbObjectError + 2450, , _
                "Missing or leaked identity properties in generated group: " & key
            Set values = Dictionary()
            For Each entry In actual
                values.Add entry("scope") & "|" & entry("name"), entry
            Next entry
            For Each entry In expected
                propertyKey = entry("scope") & "|" & entry("name")
                If values.Exists(propertyKey) = False Then Err.Raise vbObjectError + 2450, , _
                    "Missing preserved property in group " & key & ": " & propertyKey
                Set found = values(propertyKey)
                If found("value") <> entry("value") Or found("type") <> entry("type") Then _
                    Err.Raise vbObjectError + 2450, , "Changed metadata in group " & key & ": " & propertyKey
                WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("file")) & "," & Csv(entry("scope")) & "," & _
                    Csv(entry("name")) & "," & Csv(group("sourcePath")) & "," & Csv(group("sourceConfig")) & "," & _
                    Csv(entry("value")) & "," & Csv(found("value")) & "," & entry("type") & ",PASS"
                count = count + 1
            Next entry
            If group("kind") <> "COLLECTION" Then
                For Each name In Array("PartNo", "Description")
                    present = False
                    For Each entry In expected
                        If StrComp(entry("name"), CStr(name), vbTextCompare) = 0 And _
                            Len(Trim$(entry("value"))) > 0 Then present = True
                    Next entry
                    If present = False Then
                        missingStatus = "SOURCE_MISSING"
                        If group("kind") = "KIT" Then missingStatus = "KIT_NUMBER_MISSING"
                        LogLine "IDENTITY METADATA MISSING: " & key & "; " & name & "; " & group("sourcePath")
                        WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("file")) & ",effective," & Csv(CStr(name)) & "," & _
                            Csv(group("sourcePath")) & "," & Csv(group("sourceConfig")) & ",,,," & missingStatus
                    End If
                Next name
            End If
        End If
    Next key
    SaveCsvReport report, reportPrefix & "metadata.csv"
    LogLine "Metadata preservation PASS: " & count & _
        " properties retain scope, type and resolved value after reopen; kits use their own provenance identity."
End Sub

Private Sub AddTextMetadata(ByVal metadata As Collection, ByVal name As String, ByVal value As String)
    Dim entry As Object
    Set entry = Dictionary()
    entry.Add "scope", ""
    entry.Add "name", name
    entry.Add "type", 30&
    entry.Add "value", value
    metadata.Add entry
End Sub

Private Function KitMetadata(ByVal provenance As String, ByVal categoryName As String, _
    ByVal referenceBacked As Boolean) As Collection
    Dim metadata As New Collection
    Dim pattern As Object
    Dim matches As Object
    Dim tail As String
    Dim partNo As String
    Dim description As String
    tail = Mid$(provenance, InStrRev(provenance, "/") + 1)
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = True
    pattern.Pattern = "\s+"
    tail = Trim$(pattern.Replace(tail, " "))
    If referenceBacked Then
        pattern.Global = False
        pattern.Pattern = "-[0-9]+$"
        tail = pattern.Replace(tail, "")
        pattern.Pattern = "^([0-9]{3,4} [0-9]{3,4} [0-9]{2}) +(.+)$"
        Set matches = pattern.Execute(tail)
        If matches.Count = 1 Then
            partNo = matches(0).SubMatches(0)
            description = Trim$(matches(0).SubMatches(1))
        Else
            pattern.Pattern = "^[0-9]"
            If Len(tail) = 0 Or pattern.Test(tail) Then Err.Raise vbObjectError + 2454, , _
                "Cannot split reference kit provenance into PartNo and Description: " & provenance
            description = tail
        End If
    Else
        description = Trim$(categoryName) & " KIT"
    End If
    AddTextMetadata metadata, "PartNo", partNo
    AddTextMetadata metadata, "Description", description
    Set KitMetadata = metadata
End Function

Private Sub PrepareKitMetadata()
    Dim key As Variant
    Dim group As Object
    Dim metadata As Collection
    Dim entry As Object
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" And group("count") > 0 Then
            Set group("metadata") = KitMetadata(CStr(group("provenance")), CStr(group("category")), _
                CBool(Left$(CStr(key), 1) = "K"))
            Set metadata = group("metadata")
            Set entry = metadata(1)
            If Len(entry("value")) = 0 Then LogLine _
                "KIT PART NUMBER MISSING: " & key & "; text-only Description; no service number invented."
        ElseIf group("kind") = "COLLECTION" Then
            Set metadata = New Collection
            AddTextMetadata metadata, "PartNo", mFS.GetBaseName(mSourcePath) & " COMPLETE KIT"
            AddTextMetadata metadata, "Description", "Complete kit collection"
            Set group("metadata") = metadata
        End If
    Next key
End Sub

Public Sub TestKitifyMetadata()
    Dim metadata As Collection
    Dim entry As Object
    Set metadata = KitMetadata("A7393 0800 45 COMPLETE KIT/7372 1043 01 HLX-FH-SEAL KIT-1", "", True)
    Set entry = metadata(1)
    If entry("value") <> "7372 1043 01" Then Err.Raise vbObjectError + 2454, , "Kit PartNo regression."
    Set entry = metadata(2)
    If entry("value") <> "HLX-FH-SEAL KIT" Then Err.Raise vbObjectError + 2454, , "Kit Description regression."
    Set metadata = KitMetadata(" 7372" & vbTab & "1045  03 RDX5 IB SEAL KIT-12 ", "", True)
    Set entry = metadata(1)
    If entry("value") <> "7372 1045 03" Then Err.Raise vbObjectError + 2454, , "Kit whitespace regression."
    Set entry = metadata(2)
    If entry("value") <> "RDX5 IB SEAL KIT" Then Err.Raise vbObjectError + 2454, , "Kit occurrence suffix regression."
    Set metadata = KitMetadata("7372 1043 01 HLX-FH-SEAL KIT", "", True)
    Set entry = metadata(2)
    If entry("value") <> "HLX-FH-SEAL KIT" Then Err.Raise vbObjectError + 2454, , "Unnumbered occurrence suffix regression."
    Set metadata = KitMetadata("NOTE/category fallback; source top-level occurrence 0", "BEARING", False)
    Set entry = metadata(1)
    If entry("value") <> "" Then Err.Raise vbObjectError + 2454, , "Fallback invented a kit number."
    Set entry = metadata(2)
    If entry("value") <> "BEARING KIT" Then Err.Raise vbObjectError + 2454, , "Fallback Description regression."
    Set metadata = KitMetadata("Parent/RDX5 PERCUSSION SEAL KIT-1", "SEAL", True)
    Set entry = metadata(1)
    If entry("value") <> "" Then Err.Raise vbObjectError + 2454, , "Unnumbered reference invented a kit number."
    Set entry = metadata(2)
    If entry("value") <> "RDX5 PERCUSSION SEAL KIT" Then Err.Raise vbObjectError + 2454, , "Text-only provenance regression."
    TestInvalidKitMetadata
    Debug.Print "Kit metadata tests PASS"
End Sub

Private Sub TestInvalidKitMetadata()
    Dim metadata As Collection
    Dim number As Long
    On Error GoTo ExpectedFailure
    Set metadata = KitMetadata("7372 1043 INVALID KIT-1", "", True)
    On Error GoTo 0
    Err.Raise vbObjectError + 2454, , "Malformed reference kit provenance was accepted."
ExpectedFailure:
    number = Err.Number
    If number <> vbObjectError + 2454 Then Err.Raise number, , Err.Description
End Sub

Private Function PartKey(ByVal value As String) As String
    Dim pattern As Object
    Dim matches As Object
    Dim match As Object
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "(^|[^A-Z0-9])(BG[ ]*)?([0-9]{3})[ -]?([0-9]{3})[ -]?([0-9]{2})([^0-9]|$)"
    Set matches = pattern.Execute(UCase$(value))
    If matches.Count = 0 Then Exit Function
    Set match = matches(0)
    If Len(match.SubMatches(1)) > 0 Then PartKey = "BG"
    PartKey = PartKey & match.SubMatches(2) & match.SubMatches(3) & match.SubMatches(4)
End Function

Private Function Identity(ByVal path As String) As String
    Identity = UCase$(Replace(mFS.GetBaseName(path), "-kitified", "", 1, -1, vbTextCompare))
End Function

Private Function Category(ByVal note As String) As String
    Dim text As String
    text = UCase$(Trim$(note))
    text = Replace(Replace(Replace(text, "_", " "), "-", " "), ".", "")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    If Right$(text, 5) = " KITS" Then text = Left$(text, Len(text) - 5)
    If Right$(text, 4) = " KIT" Then text = Left$(text, Len(text) - 4)
    Select Case text
        Case "SEAL", "SEALS": text = "SEAL"
        Case "FASTENER", "FASTENERS": text = "FASTENER"
        Case "BEARING", "BEARINGS": text = "BEARING"
        Case "HYDR", "HYDRAULIC", "HYDR FITTING", "HYDR FITTINGS", _
            "HYDRAULIC FITTING", "HYDRAULIC FITTINGS": text = "HYDR FITTING"
    End Select
    Category = Trim$(text)
End Function

Private Function KitCategory(ByVal title As String) As String
    Select Case True
        Case InStr(title, "SEAL") > 0: KitCategory = "SEAL"
        Case InStr(title, "FASTENER") > 0: KitCategory = "FASTENER"
        Case InStr(title, "HYDR") > 0: KitCategory = "HYDR FITTING"
        Case InStr(title, "BEARING") > 0: KitCategory = "BEARING"
        Case Else: KitCategory = "EXPLICIT"
    End Select
End Function

Private Sub AddGroup(ByVal key As String, ByVal parent As String, ByVal kind As String, _
    ByVal categoryName As String, ByVal provenance As String)
    Dim group As Object
    Dim metadata As New Collection
    If mGroups.Exists(key) Then Exit Sub
    Set group = Dictionary()
    group.Add "parent", parent
    group.Add "kind", kind
    group.Add "category", categoryName
    group.Add "provenance", provenance
    group.Add "file", ""
    group.Add "count", 0&
    group.Add "building", False
    group.Add "nodeName", ""
    group.Add "configuration", ""
    group.Add "metadata", metadata
    group.Add "sourcePath", ""
    group.Add "sourceConfig", ""
    group.Add "colourIndex", -1&
    mGroups.Add key, group
End Sub

Private Sub DiscoverKits()
    Dim n As Object
    Dim parent As Object
    Dim pattern As Object
    Dim key As String
    Dim parentKit As String
    Dim p As Long
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "(^|[ _-])KITS?($|[ _-])"
    For Each n In mReference
        If n("excluded") = False Then
            If n("assembly") And pattern.Test(n("identity")) Then
                key = "K" & Format$(n("id"), "0000")
                parentKit = "ROOT"
                p = n("parent")
                Do While p > 0
                    Set parent = mReference(p)
                    If Len(parent("kit")) > 0 Then
                        parentKit = parent("kit")
                        Exit Do
                    End If
                    p = parent("parent")
                Loop
                AddGroup key, parentKit, "KIT", KitCategory(n("identity")), n("instance")
                n("kit") = key
            ElseIf n("parent") > 0 Then
                Set parent = mReference(n("parent"))
                n("kit") = parent("kit")
            End If
        End If
    Next n
    If mGroups.Count = 1 Then Err.Raise vbObjectError + 2413, , "No kit assemblies found in the reference."
End Sub

Private Function IsLeaf(ByVal n As Object) As Boolean
    IsLeaf = CBool(n("excluded") = False And n("assembly") = False)
End Function

Private Function AncestorScore(ByVal source As Object, ByVal reference As Object) As Long
    Dim a As Object
    Dim b As Object
    Dim p As Long
    Dim q As Long
    Dim score As Long
    p = source("parent")
    Do While p > 0
        Set a = mSource(p)
        q = reference("parent")
        Do While q > 0
            Set b = mReference(q)
            If a("identity") = b("identity") Then
                score = score + 100
                If p = source("parent") And q = reference("parent") Then score = score + 10
                Exit Do
            End If
            q = b("parent")
        Loop
        p = a("parent")
    Loop
    If source("category") = reference("category") Then score = score + 1
    If Len(reference("kit")) > 0 Then
        Set a = mGroups(reference("kit"))
        If a("category") = source("category") Then score = score + 2
    End If
    AncestorScore = score
End Function

Private Sub MatchOccurrences()
    Dim s As Object
    Dim r As Object
    Dim best As Object
    Dim score As Long
    Dim bestScore As Long
    Dim ambiguous As Boolean
    Dim pass As Long
    Dim other As Object
    Dim peer As Object
    Dim depth As Long
    Dim bestDepth As Long
    For pass = 1 To 2
    For Each s In mSource
        If IsLeaf(s) And s("match") = 0 Then
            Set best = Nothing
            bestScore = -1
            ambiguous = False
            For Each r In mReference
                If IsLeaf(r) And r("match") = 0 Then
                    If (s("identity") = r("identity") Or _
                        (Len(s("partKey")) > 0 And s("partKey") = r("partKey"))) And _
                        (s("config") = r("config") Or _
                            (Len(s("partKey")) > 0 And s("partKey") = r("partKey"))) Then
                        score = AncestorScore(s, r)
                        If s("config") = r("config") Then score = score + 10000
                        If s("identity") = r("identity") Then score = score + 1000
                        If pass = 2 And Len(r("kit")) > 0 Then
                            bestDepth = 0
                            For Each other In mSource
                                If other("match") > 0 Then
                                    Set peer = mReference(other("match"))
                                    If peer("kit") = r("kit") Then
                                        depth = CommonDepth(s, other)
                                        If depth > bestDepth Then bestDepth = depth
                                    End If
                                End If
                            Next other
                            score = score + bestDepth * 10
                        End If
                        If score > bestScore Then
                            Set best = r
                            bestScore = score
                            ambiguous = False
                        ElseIf score = bestScore Then
                            If r("kit") <> best("kit") Then ambiguous = True
                        End If
                    End If
                End If
            Next r
            If Not best Is Nothing And ambiguous = False Then
                s("match") = best("id")
                best("match") = s("id")
                s("reason") = "Reference identity/configuration and ancestry"
                If s("identity") <> best("identity") Then _
                    s("reason") = "Reference embedded part-number alias/configuration and ancestry; source geometry retained"
                If s("config") <> best("config") Then _
                    s("reason") = s("reason") & "; reference configuration differs; source configuration retained"
            ElseIf ambiguous Then
                s("reason") = "Ambiguous reference ownership; category/scope fallback"
            Else
                s("reason") = "No matching reference identity/configuration"
            End If
        End If
    Next s
    Next pass
End Sub

Private Function ReferenceDestination(ByVal id As Long) As String
    Dim n As Object
    Dim source As Object
    Dim key As String
    Dim parentKey As String
    Set n = mReference(id)
    If Len(n("kit")) = 0 Then Exit Function
    key = "K" & Format$(id, "0000")
    If mGroups.Exists(key) Then
        ReferenceDestination = key
        Exit Function
    End If
    key = "R" & Format$(id, "0000")
    parentKey = ReferenceDestination(n("parent"))
    If Len(parentKey) = 0 Then Err.Raise vbObjectError + 2414, , "Missing reference kit parent."
    AddGroup key, parentKey, "STRUCTURE", "", n("instance")
    Set source = SourceStructure(n)
    mGroups(key)("nodeName") = mFS.GetBaseName(source("path"))
    Set mGroups(key)("metadata") = source("metadata")
    mGroups(key)("sourcePath") = source("path")
    mGroups(key)("sourceConfig") = source("config")
    ReferenceDestination = key
End Function

Private Function SourceStructure(ByVal reference As Object) As Object
    Dim source As Object
    Dim candidate As Object
    Dim pass As Long
    For pass = 1 To 2
        For Each source In mSource
            If source("assembly") And source("excluded") = False Then
                If (pass = 1 And source("identity") = reference("identity")) Or _
                    (pass = 2 And Len(reference("partKey")) > 0 And source("partKey") = reference("partKey")) Then
                    If Not candidate Is Nothing Then
                        If candidate("path") <> source("path") Or candidate("config") <> source("config") Then _
                            Err.Raise vbObjectError + 2439, , _
                            "Ambiguous production metadata owner for reference structure: " & reference("instance")
                    End If
                    Set candidate = source
                End If
            End If
        Next source
        If Not candidate Is Nothing Then
            Set SourceStructure = candidate
            Exit Function
        End If
    Next pass
    Err.Raise vbObjectError + 2439, , "No production metadata owner for reference structure: " & reference("instance")
End Function

Private Function SourceDestination(ByVal id As Long) As String
    Dim n As Object
    Dim key As String
    Dim parentKey As String
    If id = 0 Then
        SourceDestination = "ROOT"
        Exit Function
    End If
    Set n = mSource(id)
    key = "A" & Format$(id, "0000")
    parentKey = SourceDestination(n("parent"))
    AddGroup key, parentKey, "STRUCTURE", "", n("instance")
    mGroups(key)("nodeName") = mFS.GetBaseName(n("path"))
    Set mGroups(key)("metadata") = n("metadata")
    mGroups(key)("sourcePath") = n("path")
    mGroups(key)("sourceConfig") = n("config")
    SourceDestination = key
End Function

Private Function CommonDepth(ByVal a As Object, ByVal b As Object) As Long
    Dim p As Long
    Dim q As Long
    Dim depth As Long
    Dim n As Object
    p = a("parent")
    Do While p > 0
        q = b("parent")
        Do While q > 0
            If p = q Then
                depth = 1
                Set n = mSource(p)
                Do While n("parent") > 0
                    depth = depth + 1
                    Set n = mSource(n("parent"))
                Loop
                CommonDepth = depth
                Exit Function
            End If
            Set n = mSource(q)
            q = n("parent")
        Loop
        Set n = mSource(p)
        p = n("parent")
    Loop
End Function

Private Function FallbackKit(ByVal source As Object) As String
    Dim other As Object
    Dim r As Object
    Dim group As Object
    Dim key As String
    Dim best As String
    Dim score As Long
    Dim bestScore As Long
    Dim ambiguous As Boolean
    Dim p As Long
    Dim rootId As Long
    For Each other In mSource
        If IsLeaf(other) And other("match") > 0 And other("category") = source("category") Then
            Set r = mReference(other("match"))
            If Len(r("kit")) > 0 Then
                Set group = mGroups(r("kit"))
                If group("category") = source("category") Then
                    score = CommonDepth(source, other)
                    If score > bestScore Then
                        best = r("kit")
                        bestScore = score
                        ambiguous = False
                    ElseIf score = bestScore And score > 0 And best <> r("kit") Then
                        ambiguous = True
                    End If
                End If
            End If
        End If
    Next other
    If bestScore > 0 And ambiguous = False Then
        FallbackKit = best
        Exit Function
    End If
    p = source("parent")
    Do While p > 0
        rootId = p
        Set other = mSource(p)
        p = other("parent")
    Loop
    ' Bearings are a whole-drifter pool, including parts spread across branches.
    If source("category") = "BEARING" Then rootId = 0
    key = "G" & CStr(rootId) & "_" & source("category")
    AddGroup key, "ROOT", "KIT", source("category"), _
        "NOTE/category fallback; source top-level occurrence " & CStr(rootId)
    FallbackKit = key
End Function

Private Sub AssignDestinations()
    Dim s As Object
    Dim r As Object
    Dim group As Object
    Dim destination As String
    Dim referenceKit As String
    For Each s In mSource
        If IsLeaf(s) Then
            destination = ""
            If s("match") > 0 Then
                Set r = mReference(s("match"))
                referenceKit = r("kit")
                If Len(referenceKit) > 0 Then
                    Set group = mGroups(referenceKit)
                    If Len(s("category")) = 0 Or group("category") = s("category") Or _
                        group("category") = "EXPLICIT" Then
                        destination = ReferenceDestination(r("parent"))
                    End If
                End If
            End If
            If Len(destination) = 0 Then
                If Len(s("category")) > 0 Then
                    destination = FallbackKit(s)
                    s("reason") = s("reason") & "; pooled by normalized NOTE and source scope"
                Else
                    destination = SourceDestination(s("parent"))
                    s("reason") = s("reason") & "; retained in independent production structure"
                End If
            End If
            s("destination") = destination
            Do While Len(destination) > 0
                Set group = mGroups(destination)
                group("count") = group("count") + 1
                destination = group("parent")
            Loop
        End If
    Next s
End Sub

Private Function SafeName(ByVal value As String) As String
    Dim pattern As Object
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = True
    pattern.Pattern = "[^A-Za-z0-9_-]"
    SafeName = Left$(pattern.Replace(value, "_"), 55)
End Function

Private Sub CollectKits()
    Dim key As Variant
    Dim group As Object
    Dim total As Long
    Dim colourIndex As Long
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" And group("count") > 0 Then
            group("colourIndex") = colourIndex
            colourIndex = colourIndex + 1
            If group("parent") = "ROOT" Then total = total + group("count")
        End If
    Next key
    If total = 0 Then Exit Sub
    AddGroup "COMPLETE", "ROOT", "COLLECTION", "", "Generated complete-kit collection"
    mGroups("COMPLETE")("count") = total
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" And group("count") > 0 And group("parent") = "ROOT" Then _
            group("parent") = "COMPLETE"
    Next key
End Sub

Private Function KitAppearance(ByVal index As Long) As Variant
    Dim colours As Variant
    Dim rgb As Variant
    Dim values(0 To 8) As Double
    colours = Array(Array(1#, 0#, 1#), Array(1#, 0.4, 0.7), _
        Array(0#, 1#, 0#), Array(0#, 1#, 1#), _
        Array(0.65, 0#, 0.65), Array(0.75, 0.2, 0.45), _
        Array(0#, 0.6, 0#), Array(0#, 0.6, 0.6), _
        Array(1#, 0.55, 1#), Array(1#, 0.7, 0.85), _
        Array(0.55, 1#, 0.55), Array(0.55, 1#, 1#))
    rgb = colours(index Mod (UBound(colours) + 1))
    ' SOLIDWORKS persists RGB channels at 8-bit precision.
    values(0) = Fix(CDbl(rgb(0)) * 255#) / 255#
    values(1) = Fix(CDbl(rgb(1)) * 255#) / 255#
    values(2) = Fix(CDbl(rgb(2)) * 255#) / 255#
    values(3) = 1#
    values(4) = 1#
    values(5) = 0.2
    values(6) = 0.2
    KitAppearance = values
End Function

Private Sub AssertNoMates(ByVal model As SldWorks.ModelDoc2)
    Dim feature As SldWorks.Feature
    Dim mate As SldWorks.Feature
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        If feature.GetTypeName2 = "MateGroup" Then
            Set mate = feature.GetFirstSubFeature
            If Not mate Is Nothing Then Err.Raise vbObjectError + 2451, , _
                "Generated assembly contains mates: " & model.GetTitle & "; " & mate.Name
        End If
        Set feature = feature.GetNextFeature
    Loop
End Sub

Private Sub ValidateAppearances()
    Dim key As Variant
    Dim group As Object
    Dim model As SldWorks.ModelDoc2
    Dim expected As Variant
    Dim actual As Variant
    Dim i As Long
    Dim report As Object
    Set report = NewCsvReport()
    WriteCsvLine report, "Group,Kind,Red,Green,Blue,Status"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("count") > 0 Then
            Set model = mApp.GetOpenDocumentByName(group("file"))
            If model Is Nothing Then Err.Raise vbObjectError + 2452, , _
                "Generated appearance document is not loaded: " & key
            AssertNoMates model
            If group("kind") = "KIT" Then
                expected = KitAppearance(CLng(group("colourIndex")))
                actual = model.MaterialPropertyValues
                If Not IsArray(actual) Then Err.Raise vbObjectError + 2452, , _
                    "Generated appearance is unavailable: " & key
                For i = 0 To 8
                    If Abs(CDbl(actual(i)) - CDbl(expected(i))) > MATRIX_TOLERANCE Then _
                        Err.Raise vbObjectError + 2452, , "Generated appearance differs after reopen: " & key & _
                            "; property=" & i & "; expected=" & expected(i) & "; actual=" & actual(i)
                Next i
                WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("kind")) & "," & _
                    Csv(CStr(actual(0))) & "," & Csv(CStr(actual(1))) & "," & Csv(CStr(actual(2))) & ",PASS"
            Else
                If model.Extension.HasMaterialPropertyValues Then Err.Raise vbObjectError + 2452, , _
                    "Unexpected assembly appearance outside an individual kit: " & key
                WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("kind")) & ",,,,PASS"
            End If
        End If
    Next key
    SaveCsvReport report, "appearance.csv"
    LogLine "Appearance PASS: individual kits coloured; complete-kit collection and non-kit assemblies have no assembly colour override."
    LogLine "Mates PASS: no mates in any generated assembly; all occurrences verified fixed."
End Sub

Private Function KitFileName(ByVal key As String, ByVal group As Object) As String
    Dim name As String
    Dim pattern As Object
    If Left$(key, 1) <> "K" Then
        KitFileName = "KIT_" & mRunId & "_" & SafeName(key) & "_" & SafeName(group("category")) & ".SLDASM"
        Exit Function
    End If
    name = CStr(group("provenance"))
    name = Mid$(name, InStrRev(name, "/") + 1)
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.IgnoreCase = True
    pattern.Pattern = "[<>:""/\\|?*\x00-\x1F]|[ .]$|^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)"
    If Len(name) = 0 Or pattern.Test(name) Then Err.Raise vbObjectError + 2447, , _
        "Kit provenance does not provide a valid assembly filename: " & key & "; " & group("provenance")
    KitFileName = name & ".SLDASM"
End Function

Private Function IdentityMatrix() As Variant
    Dim data(0 To 15) As Double
    data(0) = 1#
    data(4) = 1#
    data(8) = 1#
    data(12) = 1#
    IdentityMatrix = data
End Function

Private Function BuildGroup(ByVal key As String, ByVal template As String, ByVal sourceName As String) As String
    Dim group As Object
    Dim child As Object
    Dim n As Object
    Dim childKey As Variant
    Dim model As SldWorks.ModelDoc2
    Dim assembly As SldWorks.AssemblyDoc
    Dim childModel As SldWorks.ModelDoc2
    Dim manager As SldWorks.CustomPropertyManager
    Dim childPath As String
    Dim fileName As String
    Dim errors As Long
    Dim warnings As Long
    Dim saved As Boolean
    Dim saveOptions As Long
    Set group = mGroups(key)
    mStage = "Building group " & key
    If group("count") = 0 Then Exit Function
    If Len(group("file")) > 0 Then
        BuildGroup = group("file")
        Exit Function
    End If
    If group("building") Then Err.Raise vbObjectError + 2415, , "Cycle in generated hierarchy."
    group("building") = True
    For Each childKey In mGroups.Keys
        Set child = mGroups(childKey)
        If child("parent") = key And child("count") > 0 Then
            childPath = BuildGroup(CStr(childKey), template, sourceName)
        End If
    Next childKey

    Set model = mWorking
    If model Is Nothing Then Err.Raise vbObjectError + 2416, , "The working kitified assembly is unavailable."
    Set assembly = model
    Activate model
    Set manager = model.Extension.CustomPropertyManager("")
    PutProperty manager, "Skycode.Kind", CStr(group("kind"))
    PutProperty manager, "Skycode.Category", CStr(group("category"))
    PutProperty manager, "Skycode.Provenance", CStr(group("provenance"))
    PutProperty manager, "Skycode.Generated", "Kitify; fixed snapshot; do not edit by hand"
    ApplyMetadata model, group("metadata")
    If model.Extension.HasMaterialPropertyValues Then
        If model.Extension.RemoveMaterialProperty(2, Empty) = False Then Err.Raise vbObjectError + 2452, , _
            "Could not clear previous working-document appearance: " & key
    End If
    If group("kind") = "KIT" Then model.MaterialPropertyValues = KitAppearance(CLng(group("colourIndex")))

    For Each n In mSource
        If IsLeaf(n) And n("destination") = key Then
            InsertFixed model, assembly, n("path"), n("config"), n("matrix")
        End If
    Next n
    For Each childKey In mGroups.Keys
        Set child = mGroups(childKey)
        If child("parent") = key And child("count") > 0 Then
            Set childModel = OpenGeneratedChild(CStr(child("file")), CStr(child("configuration")))
            InsertFixed model, assembly, child("file"), _
                CStr(child("configuration")), IdentityMatrix()
            SaveLoadedGeneratedDocuments
            Set childModel = Nothing
            LogLine "Closing inserted child document: " & mChildPath
            mApp.CloseDoc mChildPath
            mChildPath = ""
        End If
    Next childKey
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 2418, , _
        "Generated assembly rebuild failed: " & key
    AssertNoMates model
    If key = "ROOT" Then
        fileName = sourceName & " kitified snapshot " & mRunId & ".SLDASM"
    ElseIf group("kind") = "COLLECTION" Then
        fileName = sourceName & " COMPLETE KIT.SLDASM"
    ElseIf group("kind") = "KIT" Then
        fileName = KitFileName(key, group)
    Else
        fileName = group("nodeName") & "_" & mRunId & "_" & key & ".SLDASM"
    End If
    group("file") = mFS.BuildPath(mFolder, fileName)
    If mFS.FileExists(group("file")) Then Err.Raise vbObjectError + 2419, , "Refusing to overwrite an existing assembly."
    LogLine "Saving group copy/root: " & key
    group("configuration") = model.ConfigurationManager.ActiveConfiguration.Name
    saveOptions = 3
    If key = "ROOT" Then saveOptions = 1
    saved = model.Extension.SaveAs3(group("file"), 0, saveOptions, Nothing, Nothing, errors, warnings)
    If saved = False Or errors <> 0 Then Err.Raise vbObjectError + 2420, , _
        "Save failed: " & group("file") & "; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "SAVE WARNING " & group("file") & ": " & warnings
    group("building") = False
    BuildGroup = group("file")
    LogLine "Saved " & fileName & " (" & group("count") & " part occurrences)"
    If key <> "ROOT" Then ClearWorkingAssembly
End Function

Private Function TopLevelComponents(ByVal model As SldWorks.ModelDoc2) As Object
    Dim root As SldWorks.Component2
    Dim components As Variant
    Dim item As Variant
    Dim component As SldWorks.Component2
    Dim result As Object
    Set result = Dictionary()
    Set root = model.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    If root Is Nothing Then Err.Raise vbObjectError + 2453, , "Active configuration root is unavailable: " & model.GetPathName
    components = root.GetChildren
    If IsArray(components) Then
        For Each item In components
            Set component = item
            result.Add component.Name2, component
        Next item
    End If
    Set TopLevelComponents = result
End Function

Private Sub ShowConfiguration(ByVal model As SldWorks.ModelDoc2, ByVal name As String)
    Activate model
    If StrComp(model.ConfigurationManager.ActiveConfiguration.Name, name, vbTextCompare) = 0 Then Exit Sub
    If model.ShowConfiguration2(name) = False Then Err.Raise vbObjectError + 2453, , _
        "Could not activate configuration: " & name & "; " & model.GetPathName
    If model.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 2453, , _
        "Could not rebuild activated configuration: " & name & "; " & model.GetPathName
End Sub

Private Sub SetSuppressed(ByVal component As SldWorks.Component2, ByVal suppressed As Boolean)
    Dim status As Long
    Dim state As Long
    state = 2
    If suppressed Then state = 0
    status = component.SetSuppression2(state)
    If status <> 2 Then _
        Err.Raise vbObjectError + 2453, , "Could not set configuration suppression: " & component.Name2 & _
            "; suppressed=" & suppressed & "; status=" & status
End Sub

Private Function CombineConfigurations(ByVal source As SldWorks.ModelDoc2) As String
    Dim path As String
    Dim configuration As SldWorks.Configuration
    Dim configurations As Variant
    Dim name As Variant
    Dim components As Object
    Dim component As SldWorks.Component2
    Dim assembly As SldWorks.AssemblyDoc
    Dim childModel As SldWorks.ModelDoc2
    Dim group As Object
    Dim key As Variant
    Dim n As Object
    Dim manager As SldWorks.CustomPropertyManager
    Dim sourceConfig As String
    Dim errors As Long
    Dim warnings As Long
    sourceConfig = source.ConfigurationManager.ActiveConfiguration.Name
    path = mFS.BuildPath(mFolder, mFS.GetBaseName(mSourcePath) & " kitified generated " & mRunId & ".SLDASM")
    If mFS.FileExists(path) Then Err.Raise vbObjectError + 2419, , "Refusing to overwrite combined assembly: " & path
    CloseGenerated
    Activate source
    source.ClearSelection2 True
    If source.Extension.SaveAs3(path, 0, 3, Nothing, Nothing, errors, warnings) = False Or errors <> 0 Then _
        Err.Raise vbObjectError + 2453, , "Could not clone production root; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "PRODUCTION COPY WARNING: " & warnings
    If StrComp(source.GetPathName, mSourcePath, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Save As Copy unexpectedly changed the source document identity."
    errors = 0
    warnings = 0
    Set mWorking = mApp.OpenDoc6(path, ASSEMBLY_DOCUMENT, 1, sourceConfig, errors, warnings)
    If mWorking Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Could not open production clone; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "PRODUCTION CLONE OPEN WARNING: " & warnings
    CheckVisibleDocuments
    ShowConfiguration mWorking, sourceConfig
    Set configuration = mWorking.ConfigurationManager.ActiveConfiguration
    configurations = mWorking.GetConfigurationNames
    For Each name In configurations
        If StrComp(CStr(name), sourceConfig, vbTextCompare) <> 0 Then
            If mWorking.DeleteConfiguration2(CStr(name)) = False Then Err.Raise vbObjectError + 2453, , _
                "Could not remove unused configuration from production clone: " & name
        End If
    Next name
    configuration.Name = PRODUCTION_CONFIGURATION
    If StrComp(configuration.Name, PRODUCTION_CONFIGURATION, vbBinaryCompare) <> 0 Then _
        Err.Raise vbObjectError + 2453, , "Could not name the production configuration Default."
    ApplyMetadata mWorking, mRootMetadata
    Set components = TopLevelComponents(mWorking)
    Set mProductionRoots = Dictionary()
    For Each name In components.Keys
        Set component = components(name)
        mProductionRoots.Add CStr(name), CBool(component.IsSuppressed <> False)
    Next name

    Set configuration = mWorking.AddConfiguration3(KIT_CONFIGURATION, "Generated kit hierarchy", "", 0)
    If configuration Is Nothing Then Err.Raise vbObjectError + 2453, , "Could not create kitified configuration."
    ShowConfiguration mWorking, KIT_CONFIGURATION
    Set components = TopLevelComponents(mWorking)
    For Each name In mProductionRoots.Keys
        Set component = components(name)
        SetSuppressed component, True
    Next name
    Set assembly = mWorking
    For Each n In mSource
        If IsLeaf(n) And n("destination") = "ROOT" Then _
            InsertFixed mWorking, assembly, n("path"), n("config"), n("matrix")
    Next n
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("parent") = "ROOT" And group("count") > 0 Then
            Set childModel = OpenGeneratedChild(CStr(group("file")), CStr(group("configuration")))
            InsertFixed mWorking, assembly, group("file"), CStr(group("configuration")), IdentityMatrix()
            SaveLoadedGeneratedDocuments
            Set childModel = Nothing
            mApp.CloseDoc mChildPath
            mChildPath = ""
        End If
    Next key
    Set components = TopLevelComponents(mWorking)
    Set mKitRoots = Dictionary()
    For Each name In components.Keys
        If mProductionRoots.Exists(CStr(name)) = False Then mKitRoots.Add CStr(name), True
    Next name
    If mKitRoots.Count = 0 Then Err.Raise vbObjectError + 2453, , "No generated branches were added to the clone."
    ApplyMetadata mWorking, mRootMetadata
    Set manager = mWorking.Extension.CustomPropertyManager("")
    PutProperty manager, "Skycode.Kind", "ROOT"
    PutProperty manager, "Skycode.Generated", "Kitify; Default production / kitified fixed snapshot"
    ShowConfiguration mWorking, PRODUCTION_CONFIGURATION
    Set components = TopLevelComponents(mWorking)
    For Each name In mProductionRoots.Keys
        Set component = components(name)
        If CBool(component.IsSuppressed <> False) <> CBool(mProductionRoots(name)) Then _
            Err.Raise vbObjectError + 2453, , "Production suppression changed in Default: " & name
    Next name
    For Each name In mKitRoots.Keys
        Set component = components(name)
        SetSuppressed component, True
    Next name
    ShowConfiguration mWorking, KIT_CONFIGURATION
    Set components = TopLevelComponents(mWorking)
    For Each name In mKitRoots.Keys
        Set component = components(name)
        SetSuppressed component, False
    Next name
    If mWorking.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 2453, , _
        "Combined kitified configuration rebuild failed."
    errors = 0
    warnings = 0
    If mWorking.Save3(1, errors, warnings) = False Or errors <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Could not save combined configurations; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "COMBINED SAVE WARNING: " & warnings
    mGroups("ROOT")("file") = path
    mGroups("ROOT")("configuration") = KIT_CONFIGURATION
    LogLine "Combined output: " & path
    LogLine "Default retains the source active production configuration, mates, features and original component references."
    CombineConfigurations = path
End Function

Private Sub ValidateConfigurations()
    Dim report As Object
    Dim name As Variant
    Dim config As Variant
    Dim components As Object
    Dim component As SldWorks.Component2
    Dim expected As Boolean
    Dim actual As Collection
    Dim production As Object
    Dim n As Object
    Dim s As Object
    Dim rootName As String
    Dim count As Long
    Set report = NewCsvReport()
    WriteCsvLine report, "Configuration,Branch,Occurrence,Suppressed,Status"
    For Each config In Array(PRODUCTION_CONFIGURATION, KIT_CONFIGURATION)
        ShowConfiguration mWorking, CStr(config)
        Set components = TopLevelComponents(mWorking)
        If components.Count <> mProductionRoots.Count + mKitRoots.Count Then _
            Err.Raise vbObjectError + 2453, , "Unexpected combined top-level occurrence count: " & config
        For Each name In components.Keys
            Set component = components(name)
            If mProductionRoots.Exists(CStr(name)) Then
                expected = True
                If config = PRODUCTION_CONFIGURATION Then expected = mProductionRoots(name)
                rootName = "production"
            ElseIf mKitRoots.Exists(CStr(name)) Then
                expected = CBool(config = PRODUCTION_CONFIGURATION)
                rootName = "kits"
            Else
                Err.Raise vbObjectError + 2453, , "Unexpected combined occurrence: " & name
            End If
            If CBool(component.IsSuppressed <> False) <> expected Then Err.Raise vbObjectError + 2453, , _
                "Reopened configuration suppression mismatch: " & config & "; " & name
            WriteCsvLine report, Csv(CStr(config)) & "," & rootName & "," & Csv(CStr(name)) & "," & CStr(expected) & ",PASS"
        Next name
        If config = PRODUCTION_CONFIGURATION Then
            mGroups("ROOT")("configuration") = PRODUCTION_CONFIGURATION
            ValidateMetadata "production-", True
            mGroups("ROOT")("configuration") = KIT_CONFIGURATION
            Set actual = Snapshot(mWorking)
            WriteTree actual, "production-tree.csv"
            Set production = Dictionary()
            For Each n In actual
                rootName = Split(CStr(n("instance")), "/")(0)
                If mProductionRoots.Exists(rootName) Then production.Add CStr(n("instance")), n
            Next n
            If production.Count <> mSource.Count Then Err.Raise vbObjectError + 2453, , _
                "Production tree occurrence count changed in the clone."
            For Each s In mSource
                If production.Exists(CStr(s("instance"))) = False Then Err.Raise vbObjectError + 2453, , _
                    "Production occurrence missing from Default: " & s("instance")
                Set n = production(s("instance"))
                If StrComp(n("path"), s("path"), vbTextCompare) <> 0 Or n("config") <> s("config") Or _
                    n("suppressed") <> s("suppressed") Or n("envelope") <> s("envelope") Or n("fixed") <> s("fixed") Then _
                    Err.Raise vbObjectError + 2453, , "Production occurrence changed in Default: " & s("instance")
                If s("excluded") = False Then
                    If MatrixError(n("matrix"), s("matrix")) > MATRIX_TOLERANCE Then _
                        Err.Raise vbObjectError + 2453, , "Production transform changed in Default: " & s("instance")
                End If
                count = count + 1
            Next s
        End If
    Next config
    SaveCsvReport report, "configurations.csv"
    LogLine "Configuration PASS: reciprocal production/kit suppression persisted; " & count & " production occurrences unchanged."
End Sub

Private Sub CheckVisibleDocuments()
    Dim doc As SldWorks.ModelDoc2
    Dim allowed As Boolean
    Dim count As Long
    Set doc = mApp.GetFirstDocument
    Do While Not doc Is Nothing
        If doc.Visible Then
            count = count + 1
            allowed = CBool(StrComp(doc.GetPathName, mSourcePath, vbTextCompare) = 0)
            If Not mWorking Is Nothing Then
                If doc Is mWorking Then allowed = True
            End If
            If allowed = False Or count > 2 Then Err.Raise vbObjectError + 2444, , _
                "Close other CAD document windows before running kitify: " & doc.GetTitle
        End If
        Set doc = doc.GetNext
    Loop
End Sub

Private Function OpenGeneratedChild(ByVal path As String, ByVal configuration As String) As SldWorks.ModelDoc2
    Dim previous As Boolean
    Dim child As SldWorks.ModelDoc2
    Dim errors As Long
    Dim warnings As Long
    Dim failureNumber As Long
    Dim failureDescription As String
    If Len(mChildPath) > 0 Then Err.Raise vbObjectError + 2445, , _
        "Close the previous generated child before loading another: " & mChildPath
    CheckVisibleDocuments
    previous = mApp.GetDocumentVisible(ASSEMBLY_DOCUMENT)
    On Error GoTo Failed
    mApp.DocumentVisible False, ASSEMBLY_DOCUMENT
    mChildPath = path
    LogLine "Loading one hidden child document: " & path
    Set child = mApp.OpenDoc6(path, ASSEMBLY_DOCUMENT, 1, configuration, errors, warnings)
    mApp.DocumentVisible previous, ASSEMBLY_DOCUMENT
    If child Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2417, , _
        "Could not load generated child: " & path & "; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "CHILD LOAD WARNING " & path & ": " & warnings
    LogLine "Loaded child dirty: " & CStr(child.GetSaveFlag)
    If child.Visible Then Err.Raise vbObjectError + 2445, , "Generated child unexpectedly has a visible window: " & path
    If mPeakGeneratedDocuments < 2 Then mPeakGeneratedDocuments = 2
    Set OpenGeneratedChild = child
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    mApp.DocumentVisible previous, ASSEMBLY_DOCUMENT
    Err.Raise failureNumber, "Skycode.OpenGeneratedChild", failureDescription
End Function

Private Sub ClearWorkingAssembly()
    Dim assembly As SldWorks.AssemblyDoc
    Dim components As Variant
    Dim item As Variant
    Dim component As SldWorks.Component2
    Set assembly = mWorking
    SaveLoadedGeneratedDocuments
    LogLine "Clearing working assembly references"
    mWorking.ClearSelection2 True
    components = assembly.GetComponents(True)
    If IsArray(components) Then
        For Each item In components
            Set component = item
            If component.Select4(True, Nothing, False) = False Then Err.Raise vbObjectError + 2446, , _
                "Could not select the working assembly contents for release."
        Next item
        If mWorking.Extension.DeleteSelection2(0) = False Then Err.Raise vbObjectError + 2446, , _
            "Could not clear the working assembly after saving its child copy."
    End If
    If assembly.GetComponentCount(True) <> 0 Then Err.Raise vbObjectError + 2446, , _
        "Working assembly still contains components from the previous child."
    CheckVisibleDocuments
End Sub

Private Sub SaveLoadedGeneratedDocuments()
    Dim key As Variant
    Dim group As Object
    Dim model As SldWorks.ModelDoc2
    Dim path As String
    Dim errors As Long
    Dim warnings As Long
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        path = group("file")
        If Len(path) > 0 Then
            If StrComp(mFS.GetParentFolderName(path), mFolder, vbTextCompare) <> 0 Then _
                Err.Raise vbObjectError + 2447, , "Refusing to save a document outside this output run: " & path
            Set model = mApp.GetOpenDocumentByName(path)
            If Not model Is Nothing Then
                If Not model Is mWorking Then
                    If model.GetSaveFlag Then
                        If model.Save3(1, errors, warnings) = False Or errors <> 0 Then _
                            Err.Raise vbObjectError + 2447, , "Could not save rebuilt generated child: " & path & _
                            "; errors=" & errors & "; warnings=" & warnings
                        If warnings <> 0 Then LogLine "GENERATED CHILD SAVE WARNING " & path & ": " & warnings
                    End If
                End If
            End If
        End If
    Next key
End Sub

Private Sub InsertFixed(ByVal model As SldWorks.ModelDoc2, ByVal assembly As SldWorks.AssemblyDoc, _
    ByVal path As String, ByVal config As String, ByVal data As Variant)
    Dim component As SldWorks.Component2
    Dim utility As SldWorks.MathUtility
    Activate model
    Set component = assembly.AddComponent5(path, 0, "", True, config, 0#, 0#, 0#)
    If component Is Nothing Then Err.Raise vbObjectError + 2421, , "Insertion failed: " & path
    If StrComp(component.ReferencedConfiguration, config, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 2422, , "Inserted configuration does not match: " & path
    model.ClearSelection2 True
    If component.Select4(False, Nothing, False) = False Then Err.Raise vbObjectError + 2423, , _
        "Could not select inserted component: " & path
    assembly.UnfixComponent
    Set utility = mApp.GetMathUtility
    mStage = "Positioning " & path
    component.Transform2 = utility.CreateTransform(data)
    model.ClearSelection2 True
    If component.Select4(False, Nothing, False) = False Then Err.Raise vbObjectError + 2423, , _
        "Could not reselect positioned component: " & path
    assembly.FixComponent
    model.ClearSelection2 True
    If component.IsFixed = False Then Err.Raise vbObjectError + 2424, , "Component was not fixed: " & path
    If MatrixError(component.Transform2.ArrayData, data) > MATRIX_TOLERANCE Then _
        Err.Raise vbObjectError + 2425, , "Inserted transform does not match source: " & path
End Sub

Private Sub CloseGenerated()
    If Not mWorking Is Nothing Then
        mApp.CloseDoc mWorking.GetTitle
        Set mWorking = Nothing
    End If
End Sub

Private Sub PutProperty(ByVal manager As SldWorks.CustomPropertyManager, ByVal name As String, ByVal value As String)
    Dim status As Long
    status = manager.Add3(name, 30, value, 2)
    If status <> 0 Then Err.Raise vbObjectError + 2426, , "Could not write generated property: " & name
End Sub

Private Sub Activate(ByVal model As SldWorks.ModelDoc2)
    Dim activated As SldWorks.ModelDoc2
    Dim errors As Long
    Set activated = mApp.ActivateDoc3(model.GetTitle, False, 0, errors)
    If activated Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2427, , _
        "Activation failed: " & model.GetTitle & "; errors=" & errors
End Sub

Private Function MatrixError(ByVal a As Variant, ByVal b As Variant) As Double
    Dim i As Long
    Dim difference As Double
    For i = 0 To 12
        difference = Abs(CDbl(a(i)) - CDbl(b(i)))
        If difference > MatrixError Then MatrixError = difference
    Next i
End Function

Private Sub ValidateOutput(ByVal path As String, Optional ByVal reportPrefix As String = "")
    Dim model As SldWorks.ModelDoc2
    Dim actual As Collection
    Dim n As Object
    Dim s As Object
    Dim parent As Object
    Dim group As Object
    Dim matched As Boolean
    Dim count As Long
    Dim maxError As Double
    Dim difference As Double
    Dim expectedPath As String
    Dim report As Object
    Set model = mApp.GetOpenDocumentByName(path)
    If model Is Nothing Then Err.Raise vbObjectError + 2428, , "Saved output is not loaded."
    Set actual = Snapshot(model)
    WriteTree actual, reportPrefix & "generated-tree.csv"
    Set report = NewCsvReport()
    WriteCsvLine report, "SourceOccurrence,GeneratedOccurrence,TransformMaxError,Status"
    For Each n In actual
        If n("excluded") Then GoTo NextGeneratedOccurrence
        If n("fixed") = False Then Err.Raise vbObjectError + 2435, , _
            "Reopened output contains an unfixed occurrence: " & n("instance")
        If n("assembly") Then
            If StrComp(mFS.GetParentFolderName(n("path")), mFolder, vbTextCompare) <> 0 Then _
                Err.Raise vbObjectError + 2429, , "Output references an external/production assembly: " & n("path")
            If MatrixError(n("matrix"), IdentityMatrix()) > MATRIX_TOLERANCE Then _
                Err.Raise vbObjectError + 2436, , "Generated assembly does not share the root frame: " & n("instance")
        ElseIf IsLeaf(n) Then
            matched = False
            For Each s In mSource
                If IsLeaf(s) And s("validated") = False Then
                    If StrComp(s("path"), n("path"), vbTextCompare) = 0 And s("config") = n("config") Then
                        Set group = mGroups(s("destination"))
                        expectedPath = group("file")
                        If n("parent") = 0 Then
                            Set parent = Nothing
                            matched = CBool(s("destination") = "ROOT")
                        Else
                            Set parent = actual(n("parent"))
                            matched = CBool(StrComp(expectedPath, parent("path"), vbTextCompare) = 0)
                        End If
                        difference = MatrixError(s("matrix"), n("matrix"))
                        If difference > MATRIX_TOLERANCE Then matched = False
                        If matched Then
                            WriteCsvLine report, Csv(s("instance")) & "," & Csv(n("instance")) & "," & _
                                Csv(CStr(difference)) & ",PASS"
                            If difference > maxError Then maxError = difference
                            n("match") = s("id")
                            s("validated") = True
                            count = count + 1
                            Exit For
                        End If
                    End If
                End If
            Next s
            If matched = False Then Err.Raise vbObjectError + 2430, , _
                "Unexpected, duplicated, misplaced or incorrectly configured output occurrence: " & n("instance")
        Else
            Err.Raise vbObjectError + 2431, , "Unexpected excluded output occurrence: " & n("instance")
        End If
NextGeneratedOccurrence:
    Next n
    For Each s In mSource
        If IsLeaf(s) And s("validated") = False Then Err.Raise vbObjectError + 2432, , _
            "Missing source occurrence: " & s("instance")
    Next s
    SaveCsvReport report, reportPrefix & "validation.csv"
    LogLine "Source conservation PASS: " & count & " occurrences; max matrix error=" & CStr(maxError)
    ValidateGroupParents actual
End Sub

Private Sub ValidateGroupParents(ByVal actual As Collection)
    Dim key As Variant
    Dim group As Object
    Dim n As Object
    Dim parent As Object
    Dim found As Boolean
    Dim expectedName As String
    Dim actualName As String
    Dim occurrenceNumber As String
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If key <> "ROOT" And group("count") > 0 Then
            found = False
            For Each n In actual
                If n("excluded") = False And StrComp(n("path"), group("file"), vbTextCompare) = 0 Then
                    If group("kind") = "KIT" And Left$(CStr(key), 1) = "K" Then
                        expectedName = mFS.GetBaseName(KitFileName(CStr(key), group)) & "-"
                        actualName = Mid$(n("instance"), InStrRev(n("instance"), "/") + 1)
                        occurrenceNumber = Mid$(actualName, Len(expectedName) + 1)
                        If StrComp(Left$(actualName, Len(expectedName)), expectedName, vbTextCompare) <> 0 Or _
                            Len(occurrenceNumber) = 0 Or occurrenceNumber Like "*[!0-9]*" Then _
                            Err.Raise vbObjectError + 2448, , "Generated kit tree name differs from provenance: " & n("instance")
                    End If
                    If group("parent") = "ROOT" Then
                        found = CBool(n("parent") = 0)
                    ElseIf n("parent") > 0 Then
                        Set parent = actual(n("parent"))
                        found = CBool(StrComp(parent("path"), mGroups(group("parent"))("file"), vbTextCompare) = 0)
                    End If
                    Exit For
                End If
            Next n
            If found = False Then Err.Raise vbObjectError + 2433, , "Generated group has the wrong parent: " & key
        End If
    Next key
    LogLine "Kit hierarchy PASS: complete-kit collection at root; nested kit relationships preserved."
    LogLine "Kit naming PASS: reference kit filenames and reopened tree names retain the provenance tail."
    LogLine "Structure hierarchy PASS: all planned assembly parents and identity frames verified."
End Sub

Private Sub WriteTree(ByVal nodes As Collection, ByVal fileName As String)
    Dim report As Object
    Dim n As Object
    Set report = NewCsvReport()
    WriteCsvLine report, "Id,ParentId,Occurrence,File,Configuration,NOTE,Category,Assembly,Suppressed,Envelope,Excluded,Fixed,RootTransform"
    For Each n In nodes
        WriteCsvLine report, CStr(n("id")) & "," & CStr(n("parent")) & "," & Csv(n("instance")) & "," & _
            Csv(n("path")) & "," & Csv(n("config")) & "," & Csv(n("note")) & "," & Csv(n("category")) & "," & _
            CStr(n("assembly")) & "," & CStr(n("suppressed")) & "," & CStr(n("envelope")) & "," & _
            CStr(n("excluded")) & "," & CStr(n("fixed")) & "," & Csv(MatrixText(n("matrix")))
    Next n
    SaveCsvReport report, fileName
End Sub

Private Sub WriteLocations()
    Dim locations As Object
    Dim n As Object
    Dim parent As Object
    Dim key As Variant
    Dim location As String
    Dim report As Object
    Set locations = Dictionary()
    For Each n In mSource
        If IsLeaf(n) And Len(n("category")) > 0 Then
            location = "(root)"
            If n("parent") > 0 Then
                Set parent = mSource(n("parent"))
                location = parent("instance")
            End If
            key = Csv(location) & "," & Csv(n("category"))
            If locations.Exists(key) = False Then locations.Add key, 0&
            locations(key) = locations(key) + 1
        End If
    Next n
    Set report = NewCsvReport()
    WriteCsvLine report, "SourceParentOccurrence,NormalizedNOTE,DirectPartCount"
    For Each key In locations.Keys
        WriteCsvLine report, key & "," & CStr(locations(key))
    Next key
    SaveCsvReport report, "source-kit-locations.csv"
    LogLine "Mapped " & mSource.Count & " source nodes; NOTE locations=" & locations.Count
End Sub

Private Sub WritePlan()
    Dim report As Object
    Dim comparison As Object
    Dim n As Object
    Dim r As Object
    Dim group As Object
    Dim key As Variant
    Set report = NewCsvReport()
    WriteCsvLine report, "SourceId,SourceOccurrence,File,Configuration,NOTE,Category,Destination,ReferenceId,Reason"
    Set comparison = NewCsvReport()
    WriteCsvLine comparison, "Difference,SourceOccurrence,ReferenceOccurrence,Detail"
    For Each n In mSource
        If IsLeaf(n) Then
            WriteCsvLine report, n("id") & "," & Csv(n("instance")) & "," & Csv(n("path")) & "," & _
                Csv(n("config")) & "," & Csv(n("note")) & "," & Csv(n("category")) & "," & _
                Csv(n("destination")) & "," & n("match") & "," & Csv(n("reason"))
            If n("match") = 0 Then
                WriteCsvLine comparison, "SOURCE_UNMATCHED," & Csv(n("instance")) & ",," & Csv(n("reason"))
            Else
                Set r = mReference(n("match"))
                If n("config") <> r("config") Then WriteCsvLine comparison, _
                    "CONFIGURATION_DIFFERENCE," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv(n("config") & " -> " & r("config") & "; grouping by explicit part-number key only; source configuration retained")
                If n("identity") <> r("identity") Then WriteCsvLine comparison, _
                    "IDENTITY_ALIAS," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv("Same embedded part-number key; source file/geometry/configuration retained")
                If n("category") <> r("category") Then WriteCsvLine comparison, _
                    "NOTE_DIFFERENCE," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv(n("note") & " -> " & r("note") & "; source NOTE retained")
                If Len(n("category")) > 0 And Len(r("kit")) = 0 Then WriteCsvLine comparison, _
                    "PROMOTED_NOTE_PART," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv("Reference leaves NOTE-tagged part outside a kit; generated output pools it")
                If Len(r("kit")) > 0 And DestinationKit(CStr(n("destination"))) <> r("kit") Then _
                    WriteCsvLine comparison, "KIT_RULE_OVERRIDE," & Csv(n("instance")) & "," & _
                    Csv(r("instance")) & "," & Csv("Source NOTE category takes precedence over reference category kit")
            End If
        End If
    Next n
    For Each r In mReference
        If IsLeaf(r) And r("match") = 0 Then WriteCsvLine comparison, _
            "REFERENCE_UNMATCHED,," & Csv(r("instance")) & "," & Csv("Not copied: no matched source occurrence")
    Next r
    SaveCsvReport report, "membership.csv"
    SaveCsvReport comparison, "comparison.csv"
    Set report = NewCsvReport()
    WriteCsvLine report, "Group,Parent,Kind,Category,PartCount,Provenance"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("parent")) & "," & Csv(group("kind")) & "," & _
            Csv(group("category")) & "," & group("count") & "," & Csv(group("provenance"))
    Next key
    SaveCsvReport report, "groups.csv"
    WriteReferenceKitComparison
End Sub

Private Function DestinationKit(ByVal destination As String) As String
    Dim group As Object
    Do While Len(destination) > 0
        Set group = mGroups(destination)
        If group("kind") = "KIT" Then
            DestinationKit = destination
            Exit Function
        End If
        destination = group("parent")
    Loop
End Function

Private Sub WriteReferenceKitComparison()
    Dim report As Object
    Dim key As Variant
    Dim group As Object
    Dim n As Object
    Dim parent As Object
    Dim expected As Long
    Dim p As Long
    Dim referenceId As Long
    Set report = NewCsvReport()
    WriteCsvLine report, "Group,ReferenceKit,ReferenceDescendantParts,PlannedSourceDescendantParts,CountDifference,GeneratedParent"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" And Left$(key, 1) = "K" Then
            referenceId = CLng(Mid$(key, 2))
            expected = 0
            For Each n In mReference
                If IsLeaf(n) Then
                    p = n("parent")
                    Do While p > 0
                        If p = referenceId Then
                            expected = expected + 1
                            Exit Do
                        End If
                        Set parent = mReference(p)
                        p = parent("parent")
                    Loop
                End If
            Next n
            WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("provenance")) & "," & _
                CStr(expected) & "," & CStr(group("count")) & "," & _
                CStr(group("count") - expected) & "," & Csv(group("parent"))
        End If
    Next key
    SaveCsvReport report, "reference-kit-comparison.csv"
End Sub

Private Function NewCsvReport() As Object
    Dim report As Object
    Set report = CreateObject("ADODB.Stream")
    report.Type = 2
    report.Charset = "utf-8"
    report.LineSeparator = -1
    report.Open
    Set NewCsvReport = report
End Function

Private Sub WriteCsvLine(ByVal report As Object, ByVal value As String)
    report.WriteText value, 1
End Sub

Private Sub SaveCsvReport(ByVal report As Object, ByVal fileName As String)
    report.SaveToFile mFS.BuildPath(mFolder, fileName), 1
    report.Close
End Sub

Private Function Csv(ByVal value As String) As String
    Csv = """" & Replace(value, """", """""") & """"
End Function

Private Function MatrixText(ByVal data As Variant) As String
    Dim i As Long
    For i = 0 To 15
        If i > 0 Then MatrixText = MatrixText & ";"
        MatrixText = MatrixText & CStr(data(i))
    Next i
End Function

Private Sub LogLine(ByVal text As String)
    Debug.Print text
    If Not mLog Is Nothing Then mLog.WriteLine text
End Sub
