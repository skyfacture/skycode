Option Explicit

' Requires the installed SOLIDWORKS type library; no constants-library reference.
Private Const ASSEMBLY_DOCUMENT As Long = 2
Private Const DEFAULT_ASSEMBLY_TEMPLATE As Long = 9
Private Const MATRIX_TOLERANCE As Double = 0.00000001

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
    ValidateOutput result
    ValidateMetadata
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

Private Function IsIdentityProperty(ByVal name As String) As Boolean
    Select Case UCase$(name)
        Case "PARTNO", "DESCRIPTION", "GLOBAL PART NUMBER": IsIdentityProperty = True
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
                If IsIdentityProperty(CStr(name)) Then
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
                If IsIdentityProperty(CStr(name)) Then
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

Private Sub ValidateMetadata()
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
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, "metadata.csv"), False, True)
    report.WriteLine "Group,GeneratedAssembly,Scope,Property,SourceFile,SourceConfiguration,ExpectedValue,ActualValue,Type,Status"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("count") > 0 Then
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
                report.WriteLine Csv(CStr(key)) & "," & Csv(group("file")) & "," & Csv(entry("scope")) & "," & _
                    Csv(entry("name")) & "," & Csv(group("sourcePath")) & "," & Csv(group("sourceConfig")) & "," & _
                    Csv(entry("value")) & "," & Csv(found("value")) & "," & entry("type") & ",PASS"
                count = count + 1
            Next entry
            If group("kind") <> "KIT" Then
                For Each name In Array("PartNo", "Description")
                    present = False
                    For Each entry In expected
                        If StrComp(entry("name"), CStr(name), vbTextCompare) = 0 And _
                            Len(Trim$(entry("value"))) > 0 Then present = True
                    Next entry
                    If present = False Then
                        LogLine "SOURCE METADATA MISSING: " & key & "; " & name & "; " & group("sourcePath")
                        report.WriteLine Csv(CStr(key)) & "," & Csv(group("file")) & ",effective," & Csv(CStr(name)) & "," & _
                            Csv(group("sourcePath")) & "," & Csv(group("sourceConfig")) & ",,,,SOURCE_MISSING"
                    End If
                Next name
            End If
        End If
    Next key
    report.Close
    LogLine "Metadata preservation PASS: " & count & _
        " source properties retain scope, type and resolved value after reopen; no production identity leaked into kits."
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
    If key = "ROOT" Then
        fileName = sourceName & " kitified generated " & mRunId & ".SLDASM"
    ElseIf group("kind") = "KIT" Then
        fileName = "KIT_" & mRunId & "_" & SafeName(key) & "_" & SafeName(group("category")) & ".SLDASM"
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

Private Sub ValidateOutput(ByVal path As String)
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
    WriteTree actual, "generated-tree.csv"
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, "validation.csv"), False, True)
    report.WriteLine "SourceOccurrence,GeneratedOccurrence,TransformMaxError,Status"
    For Each n In actual
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
                            report.WriteLine Csv(s("instance")) & "," & Csv(n("instance")) & "," & _
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
    Next n
    For Each s In mSource
        If IsLeaf(s) And s("validated") = False Then Err.Raise vbObjectError + 2432, , _
            "Missing source occurrence: " & s("instance")
    Next s
    report.Close
    LogLine "Source conservation PASS: " & count & " occurrences; max matrix error=" & CStr(maxError)
    ValidateGroupParents actual
End Sub

Private Sub ValidateGroupParents(ByVal actual As Collection)
    Dim key As Variant
    Dim group As Object
    Dim n As Object
    Dim parent As Object
    Dim found As Boolean
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If key <> "ROOT" And group("count") > 0 Then
            found = False
            For Each n In actual
                If StrComp(n("path"), group("file"), vbTextCompare) = 0 Then
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
    LogLine "Kit hierarchy PASS: topmost kits at root; nested kit relationships preserved."
    LogLine "Structure hierarchy PASS: all planned assembly parents and identity frames verified."
End Sub

Private Sub WriteTree(ByVal nodes As Collection, ByVal fileName As String)
    Dim report As Object
    Dim n As Object
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, fileName), False, True)
    report.WriteLine "Id,ParentId,Occurrence,File,Configuration,NOTE,Category,Assembly,Suppressed,Envelope,Excluded,Fixed,RootTransform"
    For Each n In nodes
        report.WriteLine CStr(n("id")) & "," & CStr(n("parent")) & "," & Csv(n("instance")) & "," & _
            Csv(n("path")) & "," & Csv(n("config")) & "," & Csv(n("note")) & "," & Csv(n("category")) & "," & _
            CStr(n("assembly")) & "," & CStr(n("suppressed")) & "," & CStr(n("envelope")) & "," & _
            CStr(n("excluded")) & "," & CStr(n("fixed")) & "," & Csv(MatrixText(n("matrix")))
    Next n
    report.Close
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
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, "source-kit-locations.csv"), False, True)
    report.WriteLine "SourceParentOccurrence,NormalizedNOTE,DirectPartCount"
    For Each key In locations.Keys
        report.WriteLine key & "," & CStr(locations(key))
    Next key
    report.Close
    LogLine "Mapped " & mSource.Count & " source nodes; NOTE locations=" & locations.Count
End Sub

Private Sub WritePlan()
    Dim report As Object
    Dim comparison As Object
    Dim n As Object
    Dim r As Object
    Dim group As Object
    Dim key As Variant
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, "membership.csv"), False, True)
    report.WriteLine "SourceId,SourceOccurrence,File,Configuration,NOTE,Category,Destination,ReferenceId,Reason"
    Set comparison = mFS.CreateTextFile(mFS.BuildPath(mFolder, "comparison.csv"), False, True)
    comparison.WriteLine "Difference,SourceOccurrence,ReferenceOccurrence,Detail"
    For Each n In mSource
        If IsLeaf(n) Then
            report.WriteLine n("id") & "," & Csv(n("instance")) & "," & Csv(n("path")) & "," & _
                Csv(n("config")) & "," & Csv(n("note")) & "," & Csv(n("category")) & "," & _
                Csv(n("destination")) & "," & n("match") & "," & Csv(n("reason"))
            If n("match") = 0 Then
                comparison.WriteLine "SOURCE_UNMATCHED," & Csv(n("instance")) & ",," & Csv(n("reason"))
            Else
                Set r = mReference(n("match"))
                If n("config") <> r("config") Then comparison.WriteLine _
                    "CONFIGURATION_DIFFERENCE," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv(n("config") & " -> " & r("config") & "; grouping by explicit part-number key only; source configuration retained")
                If n("identity") <> r("identity") Then comparison.WriteLine _
                    "IDENTITY_ALIAS," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv("Same embedded part-number key; source file/geometry/configuration retained")
                If n("category") <> r("category") Then comparison.WriteLine _
                    "NOTE_DIFFERENCE," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv(n("note") & " -> " & r("note") & "; source NOTE retained")
                If Len(n("category")) > 0 And Len(r("kit")) = 0 Then comparison.WriteLine _
                    "PROMOTED_NOTE_PART," & Csv(n("instance")) & "," & Csv(r("instance")) & "," & _
                    Csv("Reference leaves NOTE-tagged part outside a kit; generated output pools it")
                If Len(r("kit")) > 0 And DestinationKit(CStr(n("destination"))) <> r("kit") Then _
                    comparison.WriteLine "KIT_RULE_OVERRIDE," & Csv(n("instance")) & "," & _
                    Csv(r("instance")) & "," & Csv("Source NOTE category takes precedence over reference category kit")
            End If
        End If
    Next n
    For Each r In mReference
        If IsLeaf(r) And r("match") = 0 Then comparison.WriteLine _
            "REFERENCE_UNMATCHED,," & Csv(r("instance")) & "," & Csv("Not copied: no matched source occurrence")
    Next r
    report.Close
    comparison.Close
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, "groups.csv"), False, True)
    report.WriteLine "Group,Parent,Kind,Category,PartCount,Provenance"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        report.WriteLine Csv(CStr(key)) & "," & Csv(group("parent")) & "," & Csv(group("kind")) & "," & _
            Csv(group("category")) & "," & group("count") & "," & Csv(group("provenance"))
    Next key
    report.Close
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
    Set report = mFS.CreateTextFile(mFS.BuildPath(mFolder, "reference-kit-comparison.csv"), False, True)
    report.WriteLine "Group,ReferenceKit,ReferenceDescendantParts,PlannedSourceDescendantParts,CountDifference,GeneratedParent"
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
            report.WriteLine Csv(CStr(key)) & "," & Csv(group("provenance")) & "," & _
                CStr(expected) & "," & CStr(group("count")) & "," & _
                CStr(group("count") - expected) & "," & Csv(group("parent"))
        End If
    Next key
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
