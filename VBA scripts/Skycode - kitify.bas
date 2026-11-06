Option Explicit

' ============================================================
' Skycode - Kitify Production Drifter
' Usage:
'   - Run SkycodeKitify with the saved production drifter active in SOLIDWORKS.
'   - One root configuration is generated for every kit SLDASM in the source product's PDM 12-KITS folder.
'   - A sandbox source is matched read-only to the same filename in a vault 11-ROCK DRILL CADS folder.
'   - That product's sibling 12-KITS folder supplies membership. The vault drifter is not opened or checked out.
'   - Envelope NOTE values identify matching kits and supply provenance for PartNo/Description.
'   - Call KitifyDrifter(sourcePath, outputFolder, showMessage) to identify an open source,
'     choose an output folder, or suppress the completion message.
'   - Creates verified kitified output and reports in a new folder beside the source unless outputFolder is specified.
'   - The source assembly is not saved or restructured.
' ============================================================
' Requires the installed SOLIDWORKS type library; no constants-library reference.

Private Const PART_DOCUMENT As Long = 1
Private Const ASSEMBLY_DOCUMENT As Long = 2
Private Const DEFAULT_ASSEMBLY_TEMPLATE As Long = 9
Private Const MATRIX_TOLERANCE As Double = 0.00000001
Private Const PRODUCTION_CONFIGURATION As String = "Default"
Private Const PDM_GET_MAKE_READ_ONLY As Long = 1
Private Const PDM_GET_REFERENCES As Long = 4
Private Const PDM_END_OF_LIST As Long = -2147220984
Private Const PDM_PATH_NOT_IN_VAULT As Long = -2147220949
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

Private mApp As SldWorks.SldWorks
Private mFS As Object
Private mSource As Collection
Private mPdmKits As Collection
Private mPdmVault As Object
Private mGroups As Object
Private mFolder As String
Private mRunId As String
Private mStage As String
Private mLog As Object
Private mWorking As SldWorks.ModelDoc2
Private mSourcePath As String
Private mCombinedPath As String
Private mChildPath As String
Private mPeakGeneratedDocuments As Long
Private mRootMetadata As Collection
Private mProductionRoots As Object
Private mKitRoots As Object
Private mKitGroupRoots As Object
Private mNextSuppliedId As Long

Public Sub SkycodeKitify()
    Dim result As String
    result = KitifyDrifter()
End Sub

Public Function KitifyDrifter( _
    Optional ByVal sourcePath As String = "", _
    Optional ByVal outputFolder As String = "", _
    Optional ByVal showMessage As Boolean = True) As String

    Dim source As SldWorks.ModelDoc2
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
    mStage = "Selecting the saved production assembly"
    Set source = FindSource(sourcePath)
    mSourcePath = source.GetPathName
    CheckVisibleDocuments
    mFolder = NewOutputFolder(source.GetPathName, outputFolder)
    mRunId = Format$(Now, "yyyymmdd_hhnnss") & "_" & Hex$(CLng(Timer * 1000#))
    Set mLog = mFS.CreateTextFile(mFS.BuildPath(mFolder, "run.txt"), False, True)
    LogLine "Source: " & source.GetPathName
    LogLine "Source configuration: " & source.ConfigurationManager.ActiveConfiguration.Name
    LogLine "Source initially dirty: " & CStr(source.GetSaveFlag)
    LogLine "Root kit configurations and membership: every latest SLDASM in the source product's PDM 12-KITS folder."
    LogLine "Matched source envelope NOTE values supply kit PartNo/Description provenance."

    mCombinedPath = mFS.BuildPath(mFolder, mFS.GetBaseName(mSourcePath) & _
        " kitified generated " & mRunId & ".SLDASM")
    mStage = "Creating the production-root copy before source inspection"
    If source.Extension.SaveAs3(mCombinedPath, 0, 3, Nothing, Nothing, errors, warnings) = False Or errors <> 0 Then _
        Err.Raise vbObjectError + 2453, , "Could not create production-root copy; errors=" & errors & _
        "; warnings=" & warnings
    If StrComp(source.GetPathName, mSourcePath, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Save As Copy unexpectedly changed the source document identity."
    If warnings <> 0 Then LogLine "PRODUCTION COPY WARNING: " & warnings
    LogLine "Created production-root copy before resolving or inspecting source components: " & mCombinedPath

    Set mRootMetadata = ReadMetadata(source, source.ConfigurationManager.ActiveConfiguration.Name)
    mStage = "Mapping every production occurrence before planning kits"
    Set mSource = Snapshot(source)
    WriteTree mSource, "source-tree.csv"
    WriteLocations

    mStage = "Loading latest PDM kits from the source product's 12-KITS folder"
    Set mPdmKits = LoadPdmKits(source)

    mStage = "Planning PDM kit configurations and retained production structures"
    Set mGroups = Dictionary()
    AddGroup "ROOT", "", "ROOT", "", ""
    Set mGroups("ROOT")("metadata") = mRootMetadata
    mGroups("ROOT")("sourcePath") = source.GetPathName
    mGroups("ROOT")("sourceConfig") = source.ConfigurationManager.ActiveConfiguration.Name
    DiscoverPdmKits
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
    mStage = "Verifying reopened output against the source assembly"
    ValidateOutput result, "snapshot-"
    ValidateMetadata "snapshot-"
    ValidateAppearances
    mStage = "Cloning production and combining configuration-specific trees"
    snapshotPath = result
    result = CombineConfigurations(source)
    mStage = "Reopening and verifying the production and PDM kit configurations"
    CloseGenerated
    errors = 0
    warnings = 0
    Set reopened = mApp.OpenDoc6(result, ASSEMBLY_DOCUMENT, 1, FirstKitConfiguration(), errors, warnings)
    If reopened Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Could not reopen combined output; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "COMBINED REOPEN WARNING: " & warnings
    Set mWorking = reopened
    CheckVisibleDocuments
    ValidateConfigurations
    For Each n In mSource
        n("validated") = False
    Next n
    ValidateMetadata "", True
    If Not mApp.GetOpenDocumentByName(snapshotPath) Is Nothing Then _
        Err.Raise vbObjectError + 2453, , "Construction snapshot is unexpectedly still loaded: " & snapshotPath
    mFS.DeleteFile snapshotPath, True
    LogLine "Removed verified intermediate construction snapshot: " & snapshotPath
    For Each key In mGroups.Keys
        Set n = mGroups(key)
        If key <> "ROOT" And Len(n("file")) > 0 Then mApp.CloseDoc mFS.GetFileName(n("file"))
    Next key
    LogLine "Source finally dirty: " & CStr(source.GetSaveFlag)
    LogLine "Document budget PASS: maximum 2 visible documents; peak separately opened generated documents=" & _
        CStr(mPeakGeneratedDocuments) & "; saved children closed after each insertion."
    LogLine "Child component geometry remains loaded internally by SOLIDWORKS."
    LogLine "VERIFIED: " & result
    LogLine "Source file was not saved or restructured. See membership.csv and pdm-kit-comparison.csv for PDM kit membership."
    mLog.Close
    Set mLog = Nothing
    If Not original Is Nothing Then Activate original
    KitifyDrifter = result
    If showMessage Then MsgBox "Kitified assembly verified:" & vbCrLf & result & vbCrLf & _
        "Reports: " & mFolder & vbCrLf & _
        "Review membership.csv and pdm-kit-comparison.csv for PDM kit membership.", vbInformation, "Skycode kitify"
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = mStage & ": " & Err.Description & vbCrLf & _
        "Partial outputs may remain in: " & mFolder & vbCrLf & _
        "The source assembly has not been saved by this macro."
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
    Set mPdmKits = Nothing
    Set mPdmVault = Nothing
    Set mGroups = Nothing
    Set mLog = Nothing
    Set mWorking = Nothing
    mSourcePath = ""
    mCombinedPath = ""
    mChildPath = ""
    mPeakGeneratedDocuments = 0
    Set mRootMetadata = Nothing
    Set mProductionRoots = Nothing
    Set mKitRoots = Nothing
    Set mKitGroupRoots = Nothing
    mFolder = ""
    mRunId = ""
    mNextSuppliedId = -1
End Sub

Private Function Dictionary() As Object
    Set Dictionary = CreateObject("Scripting.Dictionary")
    Dictionary.CompareMode = vbTextCompare
End Function

Private Function FindSource(ByVal path As String) As SldWorks.ModelDoc2
    Dim doc As SldWorks.ModelDoc2
    Dim found As SldWorks.ModelDoc2
    Dim active As SldWorks.ModelDoc2
    If Len(path) > 0 Then
        Set doc = mApp.GetFirstDocument
        Do While Not doc Is Nothing
            If doc.GetType = ASSEMBLY_DOCUMENT And Len(doc.GetPathName) > 0 Then
                If StrComp(doc.GetPathName, path, vbTextCompare) = 0 Then
                    Set found = doc
                    Exit Do
                End If
            End If
            Set doc = doc.GetNext
        Loop
    Else
        Set active = mApp.ActiveDoc
        If Not active Is Nothing Then
            If active.GetType = ASSEMBLY_DOCUMENT Then
                If Len(active.GetPathName) = 0 Then Err.Raise vbObjectError + 2403, , _
                    "Save the active production assembly before running Kitify."
                Set found = active
            End If
        End If
        If found Is Nothing Then
            Set doc = mApp.GetFirstDocument
            Do While Not doc Is Nothing
                If doc.GetType = ASSEMBLY_DOCUMENT And Len(doc.GetPathName) > 0 Then
                    If Not found Is Nothing Then Err.Raise vbObjectError + 2402, , _
                        "Activate the intended production assembly or pass its sourcePath."
                    Set found = doc
                End If
                Set doc = doc.GetNext
            Loop
        End If
    End If
    If found Is Nothing Then Err.Raise vbObjectError + 2403, , _
        "Open and activate the saved production assembly, or pass its sourcePath."
    Set FindSource = found
End Function

Private Function FindPdmKitsFolder(ByVal sourcePath As String) As String
    Dim ancestor As String
    Dim candidate As String
    Dim cadFolder As String
    Dim parent As String
    Dim folder As Object
    ancestor = mFS.GetParentFolderName(sourcePath)
    Do While Len(ancestor) > 0
        candidate = mFS.BuildPath(ancestor, "12-KITS")
        cadFolder = mFS.BuildPath(ancestor, "11-ROCK DRILL CADS")
        If mFS.FolderExists(candidate) And mFS.FolderExists(cadFolder) Then
            If StrComp(Left$(sourcePath, Len(cadFolder) + 1), cadFolder & "\", vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 2460, , _
                    "The candidate 12-KITS folder is not a sibling of the source's product CAD folder: " & candidate
            End If
            Set folder = mPdmVault.GetFolderFromPath(candidate)
            If folder Is Nothing Then Err.Raise vbObjectError + 2460, , _
                "The source-relative 12-KITS folder is not in the source PDM vault: " & candidate
            FindPdmKitsFolder = candidate
            Exit Function
        End If
        parent = mFS.GetParentFolderName(ancestor)
        If StrComp(parent, ancestor, vbTextCompare) = 0 Then Exit Do
        ancestor = parent
    Loop
    Err.Raise vbObjectError + 2460, , _
        "Could not find the source product's direct 12-KITS sibling folder beside 11-ROCK DRILL CADS: " & sourcePath
End Function

Private Function IsPdmPathMiss(ByVal failureNumber As Long, ByVal failureDescription As String) As Boolean
    IsPdmPathMiss = failureNumber = PDM_PATH_NOT_IN_VAULT Or _
        InStr(1, failureDescription, "key you specified was not found", vbTextCompare) > 0
End Function

Private Function ResolvePdmDrifter(ByVal sourcePath As String, ByRef vaultName As String) As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Set mPdmVault = CreateObject("ConisioLib.EdmVault")
    On Error Resume Next
    vaultName = CStr(mPdmVault.GetVaultNameFromPath(sourcePath))
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If failureNumber = 0 And Len(vaultName) > 0 Then
        mPdmVault.LoginAuto vaultName, 0
        ResolvePdmDrifter = sourcePath
        Exit Function
    End If
    If failureNumber <> 0 And IsPdmPathMiss(failureNumber, failureDescription) = False Then _
        Err.Raise failureNumber, "Skycode.ResolvePdmDrifter", failureDescription
    LogLine "Source is outside PDM; searching read-only auto-login vault views for " & mFS.GetFileName(sourcePath)
    ResolvePdmDrifter = FindCanonicalDrifter(sourcePath, vaultName)
End Function

Private Function AutoLoginVaultViews() As Collection
    Dim reg As Object
    Dim names As Variant
    Dim index As Long
    Dim vaultName As String
    Dim root As Variant
    Dim autoLogin As Variant
    Dim view As Object
    Dim views As New Collection
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    reg.EnumKey HKEY_LOCAL_MACHINE, _
        "SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Databases", names
    If IsArray(names) Then
        For index = LBound(names) To UBound(names)
            vaultName = CStr(names(index))
            root = ""
            autoLogin = 0
            reg.GetStringValue HKEY_LOCAL_MACHINE, _
                "SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Databases\" & vaultName, _
                "ShellRoot", root
            reg.GetDWORDValue HKEY_CURRENT_USER, _
                "SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Vaults\" & vaultName, _
                "AutoLogin", autoLogin
            If Len(CStr(root & "")) > 0 And autoLogin <> 0 Then
                Set view = Dictionary()
                view.Add "name", vaultName
                view.Add "root", CStr(root)
                views.Add view
            End If
        Next index
    End If
    Set AutoLoginVaultViews = views
End Function

Private Function CanonicalKitFolder(ByVal vault As Object, ByVal filePath As String, ByVal fileName As String) As String
    Dim cadFolder As String
    Dim kitsFolder As String
    Dim folder As Object
    If StrComp(mFS.GetFileName(filePath), fileName, vbTextCompare) <> 0 Then Exit Function
    cadFolder = mFS.GetParentFolderName(filePath)
    If StrComp(mFS.GetFileName(cadFolder), "11-ROCK DRILL CADS", vbTextCompare) <> 0 Then Exit Function
    kitsFolder = mFS.BuildPath(mFS.GetParentFolderName(cadFolder), "12-KITS")
    If mFS.FolderExists(kitsFolder) = False Then Exit Function
    Set folder = vault.GetFolderFromPath(kitsFolder)
    If folder Is Nothing Then Exit Function
    CanonicalKitFolder = kitsFolder
End Function

Private Function FirstSearchResult(ByVal search As Object) As Object
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error Resume Next
    Set FirstSearchResult = search.GetFirstResult
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If failureNumber <> 0 And failureNumber <> PDM_END_OF_LIST Then _
        Err.Raise failureNumber, "Skycode.FirstSearchResult", failureDescription
End Function

Private Function NextSearchResult(ByVal search As Object) As Object
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error Resume Next
    Set NextSearchResult = search.GetNextResult
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If failureNumber <> 0 And failureNumber <> PDM_END_OF_LIST Then _
        Err.Raise failureNumber, "Skycode.NextSearchResult", failureDescription
End Function

Private Function FindCanonicalDrifter(ByVal sourcePath As String, ByRef vaultName As String) As String
    Dim views As Collection
    Dim view As Object
    Dim vault As Object
    Dim search As Object
    Dim result As Object
    Dim matches As Object
    Dim rejected As String
    Dim searched As String
    Dim loginFailures As String
    Dim fileName As String
    Dim path As String
    Dim key As Variant
    Dim failureNumber As Long
    Dim failureDescription As String
    fileName = mFS.GetFileName(sourcePath)
    Set views = AutoLoginVaultViews()
    If views.Count = 0 Then Err.Raise vbObjectError + 2460, , _
        "No auto-login SOLIDWORKS PDM vault view is available for " & fileName
    Set matches = Dictionary()
    For Each view In views
        searched = searched & CStr(view("name")) & " (" & CStr(view("root")) & "); "
        Set vault = CreateObject("ConisioLib.EdmVault")
        On Error Resume Next
        vault.LoginAuto CStr(view("name")), 0
        failureNumber = Err.Number
        failureDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If failureNumber <> 0 Or CBool(vault.IsLoggedIn <> False) = False Then
            loginFailures = loginFailures & CStr(view("name")) & ": " & failureDescription & "; "
        Else
            Set search = vault.CreateSearch
            search.FileName = fileName
            Set result = FirstSearchResult(search)
            Do While Not result Is Nothing
                path = CStr(result.Path)
                If Len(CanonicalKitFolder(vault, path, fileName)) > 0 Then
                    If matches.Exists(path) = False Then matches.Add path, vault
                ElseIf Len(path) > 0 Then
                    rejected = rejected & path & "; "
                End If
                Set result = NextSearchResult(search)
            Loop
        End If
    Next view
    If matches.Count = 1 Then
        For Each key In matches.Keys
            path = CStr(key)
        Next key
        Set mPdmVault = matches(path)
        vaultName = CStr(mPdmVault.GetVaultNameFromPath(path))
        LogLine "Matched read-only PDM drifter: " & path
        FindCanonicalDrifter = path
        Exit Function
    End If
    If matches.Count = 0 Then Err.Raise vbObjectError + 2460, , _
        "No read-only PDM drifter named " & fileName & _
        " is directly in 11-ROCK DRILL CADS beside 12-KITS. Searched: " & searched & _
        "Rejected: " & rejected & "Login: " & loginFailures
    path = ""
    For Each key In matches.Keys
        path = path & CStr(key) & "; "
    Next key
    Err.Raise vbObjectError + 2460, , "Multiple PDM drifters match " & fileName & ": " & path
End Function

Private Function OpenPdmAssembly(ByVal path As String) As SldWorks.ModelDoc2
    Dim previousVisible As Boolean
    Dim model As SldWorks.ModelDoc2
    Dim errors As Long
    Dim warnings As Long
    Dim failureNumber As Long
    Dim failureDescription As String
    previousVisible = mApp.GetDocumentVisible(ASSEMBLY_DOCUMENT)
    On Error GoTo Failed
    mApp.DocumentVisible False, ASSEMBLY_DOCUMENT
    Set model = mApp.OpenDoc6(path, ASSEMBLY_DOCUMENT, 3, "", errors, warnings)
    mApp.DocumentVisible previousVisible, ASSEMBLY_DOCUMENT
    If model Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2461, , _
        "Could not open latest PDM kit read-only: " & path & "; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "PDM KIT LOAD WARNING " & path & ": " & warnings
    If model.Visible Then Err.Raise vbObjectError + 2461, , _
        "PDM kit unexpectedly opened visibly: " & path
    Set OpenPdmAssembly = model
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    mApp.DocumentVisible previousVisible, ASSEMBLY_DOCUMENT
    Err.Raise failureNumber, "Skycode.OpenPdmAssembly", failureDescription
End Function

Private Function LoadPdmKits(ByVal source As SldWorks.ModelDoc2) As Collection
    Dim kits As New Collection
    Dim vaultName As String
    Dim folderPath As String
    Dim folder As Object
    Dim file As Object
    Dim kit As Object
    Dim model As SldWorks.ModelDoc2
    Dim nodes As Collection
    Dim node As Object
    Dim occurrences As Collection
    Dim members As Object
    Dim memberKey As String
    Dim position As Object
    Dim folderId As Variant
    Dim latestVersion As Variant
    Dim localVersion As Long
    Dim canonicalPath As String
    Dim path As String
    Dim fileName As String
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim entry As Object
    Dim total As Long
    canonicalPath = ResolvePdmDrifter(source.GetPathName, vaultName)
    folderPath = FindPdmKitsFolder(canonicalPath)
    Set folder = mPdmVault.GetFolderFromPath(folderPath)
    If folder Is Nothing Then Err.Raise vbObjectError + 2460, , _
        "Could not resolve the source-relative PDM kit folder: " & folderPath
    LogLine "PDM vault: " & vaultName
    LogLine "PDM canonical drifter: " & canonicalPath
    If StrComp(canonicalPath, source.GetPathName, vbTextCompare) <> 0 Then _
        LogLine "Canonical PDM drifter was not opened, saved, or checked out."
    LogLine "PDM kit folder: " & folderPath
    Set position = folder.GetFirstFilePosition
    Do
        Set file = Nothing
        On Error Resume Next
        Set file = folder.GetNextFile(position)
        failureNumber = Err.Number
        failureDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If failureNumber <> 0 Then
            If failureNumber = PDM_END_OF_LIST Then Exit Do
            Err.Raise failureNumber, "Skycode.LoadPdmKits", failureDescription
        End If
        fileName = CStr(file.Name)
        If LCase$(mFS.GetExtensionName(fileName)) = "sldasm" Then
            folderId = CLng(folder.ID)
            latestVersion = CLng(file.CurrentVersion)
            localVersion = CLng(file.GetLocalVersionNo(folderId))
            path = CStr(file.GetLocalPath(folderId))
            mStage = "Loading latest PDM kit " & fileName
            ' Late-bound GetFileCopy reports a file locked by the open source session as error 438.
            ' A cached latest copy is already readable; do not copy over files SOLIDWORKS has open.
            If localVersion <> CLng(latestVersion) Or Len(path) = 0 Or mFS.FileExists(path) = False Then
                LogLine "Fetching PDM kit version " & CStr(latestVersion) & " (local=" & CStr(localVersion) & "): " & fileName
                On Error Resume Next
                file.GetFileCopy 0, latestVersion, folderId, _
                    PDM_GET_MAKE_READ_ONLY Or PDM_GET_REFERENCES, ""
                failureNumber = Err.Number
                failureDescription = Err.Description
                Err.Clear
                On Error GoTo 0
                If failureNumber <> 0 Then Err.Raise vbObjectError + 2462, , _
                    "Could not fetch latest PDM kit " & fileName & ": " & failureDescription & _
                    " (" & CStr(failureNumber) & ")"
                localVersion = CLng(file.GetLocalVersionNo(folderId))
                path = CStr(file.GetLocalPath(folderId))
                If localVersion <> CLng(latestVersion) Or Len(path) = 0 Or mFS.FileExists(path) = False Then _
                    Err.Raise vbObjectError + 2462, , _
                    "PDM did not materialize the latest kit version and its missing references: " & fileName & _
                    "; expected=" & CStr(latestVersion) & "; local=" & CStr(localVersion) & "; path=" & path
            Else
                LogLine "Using cached latest PDM kit without recopying files open in SOLIDWORKS: " & fileName & _
                    "; version=" & CStr(localVersion)
            End If
            Set model = OpenPdmAssembly(path)
            Set nodes = Snapshot(model)
            Set members = Dictionary()
            Set occurrences = New Collection
            total = 0
            For Each node In nodes
                If IsLeaf(node) Then
                    memberKey = OccurrenceKey(node)
                    If Len(memberKey) = 0 Then Err.Raise vbObjectError + 2463, , _
                        "PDM kit part has no usable identity: " & node("instance")
                    If members.Exists(memberKey) = False Then members.Add memberKey, 0&
                    members(memberKey) = CLng(members(memberKey)) + 1
                    node("used") = False
                    occurrences.Add node
                    total = total + 1
                End If
            Next node
            Set kit = Dictionary()
            kit.Add "name", fileName
            kit.Add "path", path
            kit.Add "identity", Identity(path)
            kit.Add "partKey", PartKey(Identity(path))
            kit.Add "note", ReadNote(model, model.ConfigurationManager.ActiveConfiguration.Name)
            Set kit("metadata") = ReadMetadata(model, model.ConfigurationManager.ActiveConfiguration.Name)
            kit.Add "partNo", MetadataValue(kit("metadata"), "PartNo")
            kit.Add "description", MetadataValue(kit("metadata"), "Description")
            kit.Add "group", ""
            kit.Add "parent", ""
            kit.Add "sourceEnvelope", ""
            kit.Add "rootConfiguration", ""
            kit.Add "memberCount", total
            Set kit("members") = members
            Set kit("occurrences") = occurrences
            kits.Add kit
            LogLine "Read-only PDM kit: " & fileName & "; part occurrences=" & total
            mApp.CloseDoc model.GetTitle
            Set model = Nothing
        End If
    Loop
    If kits.Count = 0 Then Err.Raise vbObjectError + 2462, , _
        "No SLDASM kit assemblies were found in the source-relative PDM folder: " & folderPath
    Set LoadPdmKits = kits
End Function

Private Function Identity(ByVal path As String) As String
    Identity = UCase$(Replace(mFS.GetBaseName(path), "-kitified", "", 1, -1, vbTextCompare))
End Function

Private Function PartKey(ByVal value As String) As String
    Dim pattern As Object
    Dim matches As Object
    Dim match As Object
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "(^|[^A-Z0-9])(BG[ ]*)?([0-9]{3,4})[ -]?([0-9]{3,4})[ -]?([0-9]{2})([^0-9]|$)"
    Set matches = pattern.Execute(UCase$(value))
    If matches.Count = 0 Then Exit Function
    Set match = matches(0)
    If Len(match.SubMatches(1)) > 0 Then PartKey = "BG"
    PartKey = PartKey & match.SubMatches(2) & match.SubMatches(3) & match.SubMatches(4)
End Function

Private Function OccurrenceKey(ByVal node As Object) As String
    If Len(CStr(node("partKey"))) > 0 Then
        OccurrenceKey = "PN:" & CStr(node("partKey"))
    Else
        OccurrenceKey = "ID:" & CStr(node("identity"))
    End If
End Function

Private Function NormalizeLabel(ByVal value As String) As String
    Dim pattern As Object
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = True
    pattern.Pattern = "[^A-Z0-9]"
    NormalizeLabel = pattern.Replace(UCase$(Trim$(value)), "")
End Function

Private Function MetadataValue(ByVal metadata As Collection, ByVal propertyName As String) As String
    Dim entry As Object
    For Each entry In metadata
        If StrComp(CStr(entry("name")), propertyName, vbTextCompare) = 0 Then
            If Len(MetadataValue) = 0 Or CStr(entry("scope")) = "CONFIG" Then _
                MetadataValue = CStr(entry("value"))
        End If
    Next entry
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
        n.Add "destination", ""
        n.Add "validated", False
        n.Add "inKit", False
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
        If (key = "ROOT" Or group("count") > 0 Or group("kind") = "KIT") And _
            (rootOnly = False Or key = "ROOT") Then
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
                        If group("kind") = "KIT" Then
                            missingStatus = "PDM_KIT_METADATA_MISSING"
                            LogLine "PDM KIT METADATA MISSING: " & key & "; " & name
                        Else
                            LogLine "IDENTITY METADATA MISSING: " & key & "; " & name & "; " & group("sourcePath")
                        End If
                        WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("file")) & ",effective," & Csv(CStr(name)) & "," & _
                            Csv(group("sourcePath")) & "," & Csv(group("sourceConfig")) & ",,,," & missingStatus
                    End If
                Next name
            End If
        End If
    Next key
    SaveCsvReport report, reportPrefix & "metadata.csv"
    LogLine "Metadata preservation PASS: " & count & _
        " properties retain scope, type and resolved value after reopen; category kits have no invented service number."
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

Private Function KitMetadata(ByVal provenance As String) As Collection
    Dim metadata As New Collection
    Dim pattern As Object
    Dim matches As Object
    Dim normalized As String
    Dim partNo As String
    Dim description As String
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = True
    pattern.Pattern = "\s+"
    normalized = Trim$(pattern.Replace(provenance, " "))
    pattern.Global = False
    pattern.Pattern = "-[0-9]+$"
    normalized = pattern.Replace(normalized, "")
    pattern.Pattern = "^((BG\s*)?[0-9]{3,4}(\s+[0-9]{2,4}){1,2})\s+(.+)$"
    Set matches = pattern.Execute(normalized)
    If matches.Count <> 1 Then Err.Raise vbObjectError + 2454, , _
        "Cannot split envelope provenance into PartNo and Description: " & provenance
    partNo = Trim$(matches(0).SubMatches(0))
    description = Trim$(matches(0).SubMatches(3))
    If Len(partNo) = 0 Or Len(description) = 0 Then Err.Raise vbObjectError + 2454, , _
        "Envelope provenance has an empty PartNo or Description: " & provenance
    AddTextMetadata metadata, "PartNo", partNo
    AddTextMetadata metadata, "Description", description
    Set KitMetadata = metadata
End Function

Private Sub PrepareKitMetadata()
    Dim key As Variant
    Dim group As Object
    Dim metadata As Collection
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            If Len(CStr(group("sourceEnvelope"))) > 0 Then
                Set group("metadata") = KitMetadata(CStr(group("provenance")))
            End If
        End If
    Next key
End Sub

Public Sub TestKitifyMetadata()
    Dim metadata As Collection
    Dim entry As Object
    Set metadata = KitMetadata(" 7372" & vbTab & "1045  03 RDX5 IB SEAL KIT-12 ")
    Set entry = metadata(1)
    If entry("value") <> "7372 1045 03" Then Err.Raise vbObjectError + 2454, , "Kit PartNo split regression."
    Set entry = metadata(2)
    If entry("value") <> "RDX5 IB SEAL KIT" Then Err.Raise vbObjectError + 2454, , "Kit Description split regression."
    Set metadata = KitMetadata("BG 007 731 31 DRIFTER KIT")
    Set entry = metadata(1)
    If entry("value") <> "BG 007 731 31" Then Err.Raise vbObjectError + 2454, , "BG PartNo split regression."
    Set entry = metadata(2)
    If entry("value") <> "DRIFTER KIT" Then Err.Raise vbObjectError + 2454, , "BG Description split regression."
    TestInvalidKitMetadata
    Debug.Print "Kit metadata tests PASS"
End Sub

Private Sub TestInvalidKitMetadata()
    Dim metadata As Collection
    On Error GoTo ExpectedFailure
    Set metadata = KitMetadata("BEARING")
    Err.Raise vbObjectError + 2455, , "Expected invalid provenance to fail."
ExpectedFailure:
    If Err.Number <> vbObjectError + 2454 Then Err.Raise Err.Number, , Err.Description
End Sub

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
    group.Add "pdmPath", ""
    group.Add "pdmIdentity", ""
    group.Add "sourceEnvelope", ""
    group.Add "totalCount", 0&
    group.Add "rootConfiguration", ""
    group.Add "inKit", False
    Set group("requirements") = Dictionary()
    Set group("members") = New Collection
    Set group("validatedMembers") = Dictionary()
    mGroups.Add key, group
End Sub

Private Function IsLeaf(ByVal n As Object) As Boolean
    IsLeaf = CBool(n("excluded") = False And n("assembly") = False)
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

Private Function FindPdmKitForEnvelope(ByVal note As String) As Object
    Dim kit As Object
    Dim candidate As Object
    Dim noteKey As String
    Dim notePartKey As String
    Dim matched As Boolean
    noteKey = NormalizeLabel(note)
    notePartKey = PartKey(note)
    For Each kit In mPdmKits
        matched = False
        If Len(notePartKey) > 0 And notePartKey = kit("partKey") Then matched = True
        If noteKey = NormalizeLabel(CStr(kit("identity"))) Or _
            noteKey = NormalizeLabel(CStr(kit("partNo"))) Or _
            noteKey = NormalizeLabel(CStr(kit("description"))) Or _
            noteKey = NormalizeLabel(CStr(kit("partNo")) & CStr(kit("description"))) Or _
            noteKey = NormalizeLabel(CStr(kit("note"))) Then matched = True
        If matched Then
            If Not candidate Is Nothing Then Err.Raise vbObjectError + 2464, , _
                "Envelope NOTE matches more than one PDM kit assembly: " & note & _
                "; candidates=" & candidate("name") & ", " & kit("name")
            Set candidate = kit
        End If
    Next kit
    If candidate Is Nothing Then Err.Raise vbObjectError + 2464, , _
        "Envelope NOTE does not identify a PDM kit in the source product's 12-KITS folder: " & note
    Set FindPdmKitForEnvelope = candidate
End Function

Private Function ConfigurationName(ByVal value As String) As String
    Dim pattern As Object
    Dim result As String
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = True
    pattern.Pattern = "[\\/:*?""<>|\[\]]"
    result = Trim$(pattern.Replace(value, "_"))
    If Len(result) = 0 Then Err.Raise vbObjectError + 2464, , _
        "PDM kit filename cannot be represented as a SOLIDWORKS configuration name: " & value
    ConfigurationName = result
End Function

Private Function FirstKitConfiguration() As String
    Dim key As Variant
    Dim group As Object
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            FirstKitConfiguration = CStr(group("rootConfiguration"))
            Exit Function
        End If
    Next key
    Err.Raise vbObjectError + 2464, , "No PDM kit configurations were planned."
End Function

Private Sub DiscoverPdmKits()
    Dim n As Object
    Dim kit As Object
    Dim group As Object
    Dim key As String
    Dim configName As String
    Dim usedNames As Object
    Dim kitIndex As Long
    Set usedNames = Dictionary()
    For Each kit In mPdmKits
        kitIndex = kitIndex + 1
        key = "K" & Format$(kitIndex, "0000")
        configName = ConfigurationName(mFS.GetBaseName(CStr(kit("name"))))
        If StrComp(configName, PRODUCTION_CONFIGURATION, vbTextCompare) = 0 Or usedNames.Exists(configName) Then _
            Err.Raise vbObjectError + 2464, , _
            "PDM kit configuration name is duplicated or reserved: " & configName
        usedNames.Add configName, True
        kit("group") = key
        AddGroup key, "ROOT", "KIT", Category(CStr(kit("note"))), CStr(kit("note"))
        Set group = mGroups(key)
        Set group("requirements") = CopyDictionary(kit("members"))
        Set group("kitOccurrences") = kit("occurrences")
        group("pdmPath") = kit("path")
        group("pdmIdentity") = kit("identity")
        group("totalCount") = kit("memberCount")
        group("rootConfiguration") = configName
        Set group("metadata") = kit("metadata")
        kit("rootConfiguration") = configName
        LogLine "Planned PDM kit configuration: " & configName & " -> " & kit("name")
    Next kit

    For Each n In mSource
        If n("envelope") And n("suppressed") = False Then
            If Len(Trim$(CStr(n("note")))) = 0 Then Err.Raise vbObjectError + 2464, , _
                "Envelope has no NOTE value for PDM kit identification: " & n("instance")
            Set kit = FindPdmKitForEnvelope(CStr(n("note")))
            If Len(CStr(kit("sourceEnvelope"))) > 0 Then Err.Raise vbObjectError + 2464, , _
                "Multiple source envelopes identify the same PDM kit: " & kit("name")
            kit("sourceEnvelope") = n("instance")
            Set group = mGroups(CStr(kit("group")))
            group("sourceEnvelope") = n("instance")
            group("provenance") = CStr(n("note"))
            group("category") = Category(CStr(n("note")))
            LogLine "Envelope NOTE matched PDM kit: " & n("note") & " -> " & kit("name")
        End If
    Next n
End Sub

Private Function SupplyMissingKitMembers(ByVal group As Object, ByVal memberKey As String, ByVal needed As Long) As Long
    Dim kitNode As Object
    Dim suppliedNode As Object
    Dim supplied As Long
    If needed <= 0 Then Exit Function
    If group.Exists("kitOccurrences") = False Then Exit Function
    For Each kitNode In group("kitOccurrences")
        If supplied >= needed Then Exit For
        If kitNode("used") = False And OccurrenceKey(kitNode) = memberKey Then
            Set suppliedNode = CopyNode(kitNode)
            suppliedNode("id") = mNextSuppliedId
            mNextSuppliedId = mNextSuppliedId - 1
            suppliedNode("kitSupplied") = True
            suppliedNode("used") = True
            suppliedNode("reason") = "PDM kit occurrence is not in the active source; placed from the PDM kit"
            suppliedNode("instance") = "PDM:" & CStr(kitNode("instance"))
            group("members").Add suppliedNode
            kitNode("used") = True
            supplied = supplied + 1
            LogLine "PDM kit supplies occurrence absent from the active source: " & group("pdmIdentity") & _
                "; part=" & memberKey & "; occurrence=" & kitNode("instance")
        End If
    Next kitNode
    SupplyMissingKitMembers = supplied
End Function

Private Function CopyNode(ByVal source As Object) As Object
    Dim result As Object
    Dim key As Variant
    Set result = Dictionary()
    For Each key In source.Keys
        result.Add CStr(key), source(key)
    Next key
    Set CopyNode = result
End Function

Private Function CopyDictionary(ByVal source As Object) As Object
    Dim result As Object
    Dim key As Variant
    Set result = Dictionary()
    For Each key In source.Keys
        result.Add CStr(key), CLng(source(key))
    Next key
    Set CopyDictionary = result
End Function

Private Sub AssignDestinations()
    AssignPdmDestinations
    Exit Sub
End Sub

Private Sub AssignPdmDestinations()
    Dim s As Object
    Dim group As Object
    Dim key As Variant
    Dim memberKey As Variant
    Dim destination As String
    Dim required As Long
    Dim assigned As Long
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            For Each memberKey In group("requirements").Keys
                required = CLng(group("requirements")(CStr(memberKey)))
                assigned = 0
                For Each s In mSource
                    If IsLeaf(s) And OccurrenceKey(s) = CStr(memberKey) And assigned < required Then
                        group("members").Add s
                        s("inKit") = True
                        If Len(CStr(s("destination"))) = 0 Then s("destination") = CStr(key)
                        assigned = assigned + 1
                    End If
                Next s
                If assigned <> required Then assigned = assigned + SupplyMissingKitMembers(group, CStr(memberKey), required - assigned)
                If assigned <> required Then Err.Raise vbObjectError + 2466, , _
                    "PDM kit cannot be reproduced from the active source or the PDM kit: " & group("pdmIdentity") & _
                    "; part=" & memberKey & "; required=" & required & "; available=" & assigned
            Next memberKey
            group("count") = group("members").Count
            group("parent") = "ROOT"
            mGroups("ROOT")("count") = CLng(mGroups("ROOT")("count")) + CLng(group("count"))
        End If
    Next key
    For Each s In mSource
        If IsLeaf(s) And s("inKit") = False Then
            destination = SourceDestination(CLng(s("parent")))
            s("destination") = destination
            s("reason") = "Not required by a PDM kit; retained in the production structure"
            Do While Len(destination) > 0
                Set group = mGroups(destination)
                group("count") = CLng(group("count")) + 1
                destination = CStr(group("parent"))
            Loop
        ElseIf IsLeaf(s) Then
            s("reason") = "Included in each matching PDM kit configuration"
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
    Dim colourIndex As Long
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            group("colourIndex") = colourIndex
            colourIndex = colourIndex + 1
        End If
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
        If key = "ROOT" Or group("count") > 0 Or group("kind") = "KIT" Then
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
    LogLine "Appearance PASS: individual PDM kit assemblies are coloured; structures and snapshot root have no assembly colour override."
    LogLine "Mates PASS: no mates in any generated assembly; all occurrences verified fixed."
End Sub

Private Function KitFileName(ByVal key As String, ByVal group As Object) As String
    Dim name As String
    Dim pattern As Object
    name = Trim$(CStr(group("provenance")))
    name = Replace(Replace(Replace(name, vbCr, " "), vbLf, " "), vbTab, " ")
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = True
    pattern.Pattern = "[<>:""/\\|?*]"
    name = pattern.Replace(name, "_")
    pattern.Global = False
    pattern.Pattern = "[ .]+$"
    name = pattern.Replace(name, "")
    pattern.Pattern = "^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)"
    If Len(name) = 0 Or pattern.Test(name) Then Err.Raise vbObjectError + 2447, , _
        "Envelope NOTE does not provide a valid Windows kit filename: " & group("provenance")
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
    If group("count") = 0 And group("kind") <> "KIT" And key <> "ROOT" Then Exit Function
    If Len(group("file")) > 0 Then
        BuildGroup = group("file")
        Exit Function
    End If
    If group("building") Then Err.Raise vbObjectError + 2415, , "Cycle in generated hierarchy."
    group("building") = True
    For Each childKey In mGroups.Keys
        Set child = mGroups(childKey)
        If child("parent") = key And (child("count") > 0 Or child("kind") = "KIT") Then
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

    If group("kind") = "KIT" Then
        For Each n In group("members")
            InsertFixed model, assembly, n("path"), n("config"), n("matrix")
        Next n
    Else
        For Each n In mSource
            If IsLeaf(n) And n("inKit") = False And n("destination") = key Then _
                InsertFixed model, assembly, n("path"), n("config"), n("matrix")
        Next n
    End If
    For Each childKey In mGroups.Keys
        Set child = mGroups(childKey)
        If child("parent") = key And (child("count") > 0 Or child("kind") = "KIT") Then
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
    ElseIf group("kind") = "KIT" Then
        fileName = ConfigurationName(CStr(group("pdmIdentity"))) & " kitified " & mRunId & ".SLDASM"
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
    Dim kitConfig As String
    Dim kitRoot As String
    Dim componentName As Variant
    Dim matched As Boolean
    Dim errors As Long
    Dim warnings As Long
    sourceConfig = source.ConfigurationManager.ActiveConfiguration.Name
    path = mCombinedPath
    If Len(path) = 0 Or mFS.FileExists(path) = False Then Err.Raise vbObjectError + 2419, , _
        "Production-root copy is missing: " & path
    CloseGenerated
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

    Set assembly = mWorking
    For Each n In mSource
        If IsLeaf(n) And n("inKit") = False And n("destination") = "ROOT" Then _
            InsertFixed mWorking, assembly, n("path"), n("config"), n("matrix")
    Next n
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("parent") = "ROOT" And (group("count") > 0 Or group("kind") = "KIT") Then
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
    Set mKitGroupRoots = Dictionary()
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            matched = False
            For Each componentName In components.Keys
                Set component = components(componentName)
                If StrComp(component.GetPathName, CStr(group("file")), vbTextCompare) = 0 Then
                    mKitGroupRoots.Add CStr(key), CStr(componentName)
                    matched = True
                    Exit For
                End If
            Next componentName
            If matched = False Then Err.Raise vbObjectError + 2453, , _
                "Could not map PDM kit to its top-level generated branch: " & group("pdmIdentity")
        End If
    Next key
    ApplyMetadata mWorking, mRootMetadata
    Set manager = mWorking.Extension.CustomPropertyManager("")
    PutProperty manager, "Skycode.Kind", "ROOT"
    PutProperty manager, "Skycode.Generated", "Kitify; Default production / PDM kit configurations"
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
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            kitConfig = CStr(group("rootConfiguration"))
            ShowConfiguration mWorking, PRODUCTION_CONFIGURATION
            Set configuration = mWorking.AddConfiguration3(kitConfig, _
                "PDM kit: " & CStr(group("pdmIdentity")), "", 0)
            If configuration Is Nothing Then Err.Raise vbObjectError + 2453, , _
                "Could not create PDM kit configuration: " & kitConfig
            ShowConfiguration mWorking, kitConfig
            Set components = TopLevelComponents(mWorking)
            For Each name In mProductionRoots.Keys
                Set component = components(name)
                SetSuppressed component, True
            Next name
            For Each name In mKitRoots.Keys
                Set component = components(name)
                SetSuppressed component, True
            Next name
            kitRoot = CStr(mKitGroupRoots(CStr(key)))
            Set component = components(kitRoot)
            SetSuppressed component, False
            If mWorking.ForceRebuild3(False) = False Then Err.Raise vbObjectError + 2453, , _
                "PDM kit configuration rebuild failed: " & kitConfig
            LogLine "Created root kit configuration: " & kitConfig & " -> " & group("pdmIdentity")
        End If
    Next key
    ShowConfiguration mWorking, FirstKitConfiguration()
    errors = 0
    warnings = 0
    If mWorking.Save3(1, errors, warnings) = False Or errors <> 0 Then Err.Raise vbObjectError + 2453, , _
        "Could not save combined configurations; errors=" & errors & "; warnings=" & warnings
    If warnings <> 0 Then LogLine "COMBINED SAVE WARNING: " & warnings
    mGroups("ROOT")("file") = path
    mGroups("ROOT")("configuration") = PRODUCTION_CONFIGURATION
    LogLine "Combined output: " & path
    LogLine "Default retains production; each PDM kit assembly has exactly one root configuration with only that kit branch active."
    CombineConfigurations = path
End Function

Private Sub ValidateConfigurations()
    Dim report As Object
    Dim name As Variant
    Dim key As Variant
    Dim configurationNames As Variant
    Dim expectedConfigurations As Object
    Dim group As Object
    Dim config As String
    Dim selectedGroup As String
    Dim selectedRoot As String
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
    Set expectedConfigurations = Dictionary()
    expectedConfigurations.Add PRODUCTION_CONFIGURATION, True
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then expectedConfigurations.Add CStr(group("rootConfiguration")), True
    Next key
    configurationNames = mWorking.GetConfigurationNames
    If Not IsArray(configurationNames) Then Err.Raise vbObjectError + 2453, , _
        "Combined output configuration list is unavailable."
    For Each name In configurationNames
        If expectedConfigurations.Exists(CStr(name)) = False Then Err.Raise vbObjectError + 2453, , _
            "Unexpected configuration in combined output: " & name
        expectedConfigurations.Remove CStr(name)
    Next name
    If expectedConfigurations.Count <> 0 Then Err.Raise vbObjectError + 2453, , _
        "One or more PDM kit configurations are missing from the combined output."
    For Each key In Array(PRODUCTION_CONFIGURATION)
        config = CStr(key)
        ValidateRootConfiguration config, "", report
        mGroups("ROOT")("configuration") = PRODUCTION_CONFIGURATION
        ValidateMetadata "production-", True
        Set actual = Snapshot(mWorking)
        WriteTree actual, "production-tree.csv"
        Set production = Dictionary()
        For Each n In actual
            rootName = Split(CStr(n("instance")), "/")(0)
            If mProductionRoots.Exists(rootName) Then production.Add CStr(n("instance")), n
        Next n
        If production.Count <> mSource.Count Then Err.Raise vbObjectError + 2453, , _
            "Production tree occurrence count changed in Default."
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
    Next key
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            config = CStr(group("rootConfiguration"))
            ValidateRootConfiguration config, CStr(key), report
            ValidateKitConfigurationContents CStr(key), config, report
        End If
    Next key
    ShowConfiguration mWorking, FirstKitConfiguration()
    SaveCsvReport report, "configurations.csv"
    LogLine "Configuration PASS: Default production tree preserved; " & _
        CStr(mPdmKits.Count) & " root kit configurations match PDM membership; source occurrences checked=" & count
End Sub

Private Sub ValidateRootConfiguration(ByVal config As String, ByVal selectedGroup As String, _
    ByVal report As Object)
    Dim name As Variant
    Dim selectedRoot As String
    Dim generatedGroupKey As String
    Dim components As Object
    Dim component As SldWorks.Component2
    Dim expected As Boolean
    Dim rootName As String
    ShowConfiguration mWorking, config
    Set components = TopLevelComponents(mWorking)
    If components.Count <> mProductionRoots.Count + mKitRoots.Count Then _
        Err.Raise vbObjectError + 2453, , "Unexpected combined top-level occurrence count: " & config
    If Len(selectedGroup) > 0 Then selectedRoot = CStr(mKitGroupRoots(selectedGroup))
    For Each name In components.Keys
        Set component = components(name)
        If mProductionRoots.Exists(CStr(name)) Then
            expected = True
            If StrComp(config, PRODUCTION_CONFIGURATION, vbTextCompare) = 0 Then _
                expected = CBool(mProductionRoots(name))
            rootName = "production"
        ElseIf mKitRoots.Exists(CStr(name)) Then
            expected = True
            If Len(selectedGroup) > 0 Then expected = (StrComp(CStr(name), selectedRoot, vbTextCompare) <> 0)
            rootName = "generated"
            If StrComp(CStr(name), selectedRoot, vbTextCompare) = 0 Then rootName = "PDM kit"
        Else
            Err.Raise vbObjectError + 2453, , "Unexpected combined occurrence: " & name
        End If
        If CBool(component.IsSuppressed <> False) <> expected Then Err.Raise vbObjectError + 2453, , _
            "Reopened configuration suppression mismatch: " & config & "; " & name & _
            "; expected suppressed=" & CStr(expected)
        generatedGroupKey = GeneratedGroupForPath(CStr(component.GetPathName))
        If Len(generatedGroupKey) > 0 Then
            If MatrixError(component.Transform2.ArrayData, IdentityMatrix()) > MATRIX_TOLERANCE Then _
                Err.Raise vbObjectError + 2453, , _
                "Generated root branch does not retain its identity transform: " & config & "; " & name
        End If
        WriteCsvLine report, Csv(config) & "," & rootName & "," & Csv(CStr(name)) & "," & _
            CStr(expected) & ",PASS"
    Next name
End Sub

Private Sub ValidateKitConfigurationContents(ByVal groupKey As String, ByVal config As String, _
    ByVal report As Object)
    Dim group As Object
    Dim actual As Collection
    Dim n As Object
    Dim counts As Object
    Dim key As Variant
    Dim expected As Long
    Dim found As Long
    Set group = mGroups(groupKey)
    Set actual = Snapshot(mWorking)
    Set counts = Dictionary()
    For Each n In actual
        If IsLeaf(n) And n("excluded") = False Then
            key = OccurrenceKey(n)
            If counts.Exists(CStr(key)) = False Then counts.Add CStr(key), 0&
            counts(CStr(key)) = CLng(counts(CStr(key))) + 1
        End If
    Next n
    For Each key In group("requirements").Keys
        expected = CLng(group("requirements")(CStr(key)))
        found = 0
        If counts.Exists(CStr(key)) Then found = CLng(counts(CStr(key)))
        If found <> expected Then Err.Raise vbObjectError + 2453, , _
            "Root configuration differs from PDM kit membership: " & config & "; part=" & key & _
            "; PDM=" & expected & "; generated=" & found
        WriteCsvLine report, Csv(config) & ",PDM_MEMBER," & Csv(CStr(key)) & "," & CStr(found) & ",PASS"
        counts.Remove CStr(key)
    Next key
    If counts.Count > 0 Then
        For Each key In counts.Keys
            Err.Raise vbObjectError + 2453, , _
                "Unexpected part in PDM kit configuration: " & config & "; part=" & key & _
                "; generated=" & counts(key)
        Next key
    End If
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

Private Function InsertLoadedComponent(ByVal model As SldWorks.ModelDoc2, ByVal assembly As SldWorks.AssemblyDoc, _
    ByVal path As String, ByVal config As String) As SldWorks.Component2
    Dim opened As SldWorks.ModelDoc2
    Dim docType As Long
    Dim previousVisible As Boolean
    Dim errors As Long
    Dim warnings As Long
    If StrComp(Right$(path, 7), ".SLDASM", vbTextCompare) = 0 Then
        docType = ASSEMBLY_DOCUMENT
    ElseIf StrComp(Right$(path, 7), ".SLDPRT", vbTextCompare) = 0 Then
        docType = PART_DOCUMENT
    Else
        Err.Raise vbObjectError + 2421, , "Cannot load a kit-supplied component of this type: " & path
    End If
    previousVisible = mApp.GetDocumentVisible(docType)
    mApp.DocumentVisible False, docType
    Set opened = mApp.OpenDoc6(path, docType, 3, config, errors, warnings)
    mApp.DocumentVisible previousVisible, docType
    If opened Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2421, , _
        "Could not load kit-supplied component: " & path & "; errors=" & errors & "; warnings=" & warnings
    If opened.Visible Then Err.Raise vbObjectError + 2421, , _
        "Kit-supplied component unexpectedly has a visible window: " & path
    If warnings <> 0 Then LogLine "KIT COMPONENT LOAD WARNING " & path & ": " & warnings
    LogLine "Loaded read-only kit-supplied component for insertion: " & path & "; config=" & config
    Activate model
    Set InsertLoadedComponent = assembly.AddComponent5(path, 0, "", True, config, 0#, 0#, 0#)
End Function

Private Sub InsertFixed(ByVal model As SldWorks.ModelDoc2, ByVal assembly As SldWorks.AssemblyDoc, _
    ByVal path As String, ByVal config As String, ByVal data As Variant)
    Dim component As SldWorks.Component2
    Dim utility As SldWorks.MathUtility
    Activate model
    Set component = assembly.AddComponent5(path, 0, "", True, config, 0#, 0#, 0#)
    If component Is Nothing Then Set component = InsertLoadedComponent(model, assembly, path, config)
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

Private Function GeneratedGroupForPath(ByVal path As String) As String
    Dim key As Variant
    Dim group As Object
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If key <> "ROOT" And Len(CStr(group("file"))) > 0 Then
            If StrComp(CStr(group("file")), path, vbTextCompare) = 0 Then
                GeneratedGroupForPath = CStr(key)
                Exit Function
            End If
        End If
    Next key
End Function

Private Sub ValidateOutput(ByVal path As String, Optional ByVal reportPrefix As String = "")
    Dim model As SldWorks.ModelDoc2
    Dim actual As Collection
    Dim n As Object
    Dim s As Object
    Dim parent As Object
    Dim group As Object
    Dim member As Object
    Dim matched As Boolean
    Dim kitMemberBranch As Boolean
    Dim parentGroupKey As String
    Dim memberKey As String
    Dim key As Variant
    Dim count As Long
    Dim maxError As Double
    Dim difference As Double
    Dim expectedPath As String
    Dim report As Object
    Set model = mApp.GetOpenDocumentByName(path)
    If model Is Nothing Then Err.Raise vbObjectError + 2428, , "Saved output is not loaded."
    For Each s In mSource
        s("validated") = False
    Next s
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        Set group("validatedMembers") = Dictionary()
    Next key
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
            kitMemberBranch = False
            parentGroupKey = ""
            If n("parent") > 0 Then
                Set parent = actual(n("parent"))
                parentGroupKey = GeneratedGroupForPath(CStr(parent("path")))
            End If
            If Len(parentGroupKey) > 0 Then
                Set group = mGroups(parentGroupKey)
                kitMemberBranch = CBool(group("kind") = "KIT")
            End If
            If kitMemberBranch Then
                For Each member In group("members")
                    memberKey = CStr(member("id"))
                    If group("validatedMembers").Exists(memberKey) = False And _
                        StrComp(CStr(member("path")), CStr(n("path")), vbTextCompare) = 0 And _
                        StrComp(CStr(member("config")), CStr(n("config")), vbTextCompare) = 0 Then
                        difference = MatrixError(member("matrix"), n("matrix"))
                        If difference <= MATRIX_TOLERANCE Then
                            group("validatedMembers").Add memberKey, True
                            WriteCsvLine report, Csv(member("instance")) & "," & Csv(n("instance")) & "," & _
                                Csv(CStr(difference)) & ",PDM_KIT_PASS"
                            If difference > maxError Then maxError = difference
                            count = count + 1
                            matched = True
                            Exit For
                        End If
                    End If
                Next member
            Else
                For Each s In mSource
                    If IsLeaf(s) And s("inKit") = False And s("validated") = False Then
                        If StrComp(s("path"), n("path"), vbTextCompare) = 0 And s("config") = n("config") Then
                            expectedPath = ""
                            If s("destination") <> "ROOT" Then expectedPath = CStr(mGroups(s("destination"))("file"))
                            If n("parent") = 0 Then
                                matched = CBool(s("destination") = "ROOT")
                            ElseIf Len(parentGroupKey) > 0 Then
                                matched = CBool(parentGroupKey = CStr(s("destination")))
                            Else
                                matched = CBool(StrComp(expectedPath, CStr(parent("path")), vbTextCompare) = 0)
                            End If
                            difference = MatrixError(s("matrix"), n("matrix"))
                            If difference > MATRIX_TOLERANCE Then matched = False
                            If matched Then
                                WriteCsvLine report, Csv(s("instance")) & "," & Csv(n("instance")) & "," & _
                                    Csv(CStr(difference)) & ",PASS"
                                If difference > maxError Then maxError = difference
                                s("validated") = True
                                count = count + 1
                                Exit For
                            End If
                        End If
                    End If
                Next s
            End If
            If matched = False Then Err.Raise vbObjectError + 2430, , _
                "Unexpected, duplicated, misplaced or incorrectly configured output occurrence: " & n("instance")
        Else
            Err.Raise vbObjectError + 2431, , "Unexpected excluded output occurrence: " & n("instance")
        End If
NextGeneratedOccurrence:
    Next n
    For Each s In mSource
        If IsLeaf(s) And s("inKit") = False And s("validated") = False Then Err.Raise vbObjectError + 2432, , _
            "Missing source occurrence: " & s("instance")
    Next s
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            For Each member In group("members")
                If group("validatedMembers").Exists(CStr(member("id"))) = False Then Err.Raise _
                    vbObjectError + 2432, , "Missing PDM kit occurrence: " & group("pdmIdentity") & _
                    "; source occurrence=" & member("instance")
            Next member
        End If
    Next key
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
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If key <> "ROOT" And (group("count") > 0 Or group("kind") = "KIT") Then
            found = False
            For Each n In actual
                If n("excluded") = False And StrComp(n("path"), group("file"), vbTextCompare) = 0 Then
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
    LogLine "PDM kit roots PASS: every PDM kit has a separate generated top-level branch."
    LogLine "PDM kit membership PASS: overlapping source occurrences are repeated only across their respective kit branches."
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
    Dim group As Object
    Dim memberKey As Variant
    Dim membersByKey As Object
    Dim member As Object
    Dim key As Variant
    Dim status As String
    Set report = NewCsvReport()
    WriteCsvLine report, "SourceId,SourceOccurrence,File,Configuration,NOTE,Destination,MembershipConfigurations,Reason"
    Set comparison = NewCsvReport()
    WriteCsvLine comparison, "Classification,PDMKitAssembly,Configuration,PartIdentity,PDMMemberCount,SourceAssignments,Status"
    For Each n In mSource
        If IsLeaf(n) Then
            WriteCsvLine report, n("id") & "," & Csv(n("instance")) & "," & Csv(n("path")) & "," & _
                Csv(n("config")) & "," & Csv(n("note")) & "," & Csv(n("destination")) & "," & _
                Csv(SourceKitMemberships(n)) & "," & Csv(n("reason"))
        End If
    Next n
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            For Each member In group("members")
                If member.Exists("kitSupplied") Then
                    If CBool(member("kitSupplied")) Then _
                        WriteCsvLine report, member("id") & "," & Csv(member("instance")) & "," & Csv(member("path")) & "," & _
                            Csv(member("config")) & "," & Csv("") & "," & Csv(CStr(key)) & "," & _
                            Csv(group("rootConfiguration")) & "," & Csv(member("reason"))
                End If
            Next member
        End If
    Next key
    SaveCsvReport report, "membership.csv"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            Set membersByKey = Dictionary()
            For Each member In group("members")
                memberKey = OccurrenceKey(member)
                If membersByKey.Exists(CStr(memberKey)) = False Then membersByKey.Add CStr(memberKey), 0&
                membersByKey(CStr(memberKey)) = CLng(membersByKey(CStr(memberKey))) + 1
            Next member
            For Each memberKey In group("requirements").Keys
                status = "PASS"
                For Each member In group("members")
                    If OccurrenceKey(member) = CStr(memberKey) Then
                        If member.Exists("kitSupplied") Then
                            If CBool(member("kitSupplied")) Then status = "KIT_SUPPLIED"
                        End If
                    End If
                Next member
                WriteCsvLine comparison, "PDM_KIT," & Csv(group("pdmIdentity")) & "," & _
                    Csv(group("rootConfiguration")) & "," & Csv(CStr(memberKey)) & "," & _
                    CStr(group("requirements")(CStr(memberKey))) & "," & CStr(membersByKey(CStr(memberKey))) & "," & status
            Next memberKey
        End If
    Next key
    SaveCsvReport comparison, "pdm-kit-comparison.csv"
    Set report = NewCsvReport()
    WriteCsvLine report, "Group,Parent,Kind,PDMKitAssembly,Configuration,PartCount,Provenance"
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        WriteCsvLine report, Csv(CStr(key)) & "," & Csv(group("parent")) & "," & Csv(group("kind")) & "," & _
            Csv(group("pdmIdentity")) & "," & Csv(group("rootConfiguration")) & "," & _
            group("count") & "," & Csv(group("provenance"))
    Next key
    SaveCsvReport report, "groups.csv"
End Sub

Private Function SourceKitMemberships(ByVal sourceNode As Object) As String
    Dim key As Variant
    Dim group As Object
    Dim member As Object
    For Each key In mGroups.Keys
        Set group = mGroups(key)
        If group("kind") = "KIT" Then
            For Each member In group("members")
                If CLng(member("id")) = CLng(sourceNode("id")) Then
                    If Len(SourceKitMemberships) > 0 Then SourceKitMemberships = SourceKitMemberships & ";"
                    SourceKitMemberships = SourceKitMemberships & CStr(group("rootConfiguration"))
                    Exit For
                End If
            Next member
        End If
    Next key
End Function

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
