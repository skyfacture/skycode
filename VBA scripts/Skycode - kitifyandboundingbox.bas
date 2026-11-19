Attribute VB_Name = "modKitifyAndBoundingBox"
Option Explicit

' Kitify finishes first. The bounding box is added only after every kit
' configuration exists, and only while Default is active. A box added to the
' root before AddConfiguration3 is copied, resolved, into every kit configuration.
' InsertEnvelope still resolves the new envelope in existing configurations, and
' deleting it from one configuration deletes it from Default. Non-Default
' configurations are therefore confined by suppression after the insert.
' The 2026-11-19 Composer import did not see that suppressed box in kit configurations.
' Requires Skycode - kitify.bas and Skycode - sldworks bounding box.bas.

Private Const ASSEMBLY_DOCUMENT As Long = 2
Private Const SW_TOGGLE_EXTREF_NO_PROMPT_OR_SAVE As Long = 15
Private Const SW_TOGGLE_WARN_SAVE_UPDATE_ERRORS As Long = 329
Private Const SW_TOGGLE_WARN_SAVING_REFERENCED_DOC As Long = 334
Private Const SW_TOGGLE_AUTO_DISMISS_OPEN_MESSAGES As Long = 564
Private Const PRODUCTION_CONFIGURATION As String = "Default"
Private Const BOX_FEATURE_NAME As String = "Skycode_BoundingBox"

#If VBA7 Then
Private Declare PtrSafe Function CoRegisterMessageFilter Lib "ole32" ( _
    ByVal lpMessageFilter As LongPtr, ByRef lplpMessageFilter As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private mPreviousMessageFilter As LongPtr
#Else
Private Declare Function CoRegisterMessageFilter Lib "ole32" ( _
    ByVal lpMessageFilter As Long, ByRef lplpMessageFilter As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private mPreviousMessageFilter As Long
#End If
Private mFilterRemoved As Boolean
Private mDeclinerPid As Double
Private mDeclinerScript As String

Public Sub KitifyAndBoundingBox()
    Dim result As String
    result = RunKitifyAndBoundingBox()
    Debug.Print "Kitify and bounding box complete: " & result
End Sub

Public Function AddBoundingBoxToKitifiedAssembly(Optional ByVal combined As String = "") As String
    Dim swApp As SldWorks.SldWorks
    Dim source As SldWorks.ModelDoc2
    Dim fs As Object
    Dim sourcePath As String
    Dim sourceSize As Double
    Dim sourceTime As Date
    Dim stage As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed

    stage = "Connecting to the running SOLIDWORKS session"
    Set swApp = AttachSolidWorks()
    If swApp Is Nothing Then Err.Raise vbObjectError + 2500, , _
        "SOLIDWORKS is not running in this session. Nothing was changed."
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(combined) = 0 Then
        Set source = swApp.ActiveDoc
        If source Is Nothing Then Err.Raise vbObjectError + 2520, , _
            "Pass the kitified assembly path, or activate it before adding the bounding box."
        combined = source.GetPathName
    End If
    If fs.FileExists(combined) = False Then Err.Raise vbObjectError + 2521, , _
        "Kitified assembly does not exist: " & combined
    sourcePath = SourceBesideOutput(fs, combined)
    If Len(sourcePath) = 0 Or fs.FileExists(sourcePath) = False Then _
        Err.Raise vbObjectError + 2522, , "Could not find the source assembly beside " & combined
    sourceSize = fs.GetFile(sourcePath).Size
    sourceTime = fs.GetFile(sourcePath).DateLastModified
    RemoveMessageFilter swApp
    StartDecliner swApp, fs, "C:\skycode\logs\kitifyandboundingbox-declined.txt"
    AddBoundingBoxToKitifiedAssembly = FinishWithBoundingBox(swApp, fs, sourcePath, sourceSize, sourceTime, combined, stage)
    StopDecliner
    RestoreMessageFilter
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = stage & ": " & Err.Description
    Debug.Print failureDescription
    WriteFailure failureDescription
    On Error Resume Next
    StopDecliner
    RestoreMessageFilter
    On Error GoTo 0
    Err.Raise failureNumber, "KitifyAndBoundingBox", failureDescription
End Function

Public Function RunKitifyAndBoundingBox() As String
    Dim swApp As SldWorks.SldWorks
    Dim source As SldWorks.ModelDoc2
    Dim generated As SldWorks.ModelDoc2
    Dim fs As Object
    Dim sourcePath As String
    Dim sourceSize As Double
    Dim sourceTime As Date
    Dim combined As String
    Dim report As String
    Dim errors As Long
    Dim warnings As Long
    Dim stage As String
    Dim savedToggles As Object
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed

    stage = "Connecting to the running SOLIDWORKS session"
    Set swApp = AttachSolidWorks()
    If swApp Is Nothing Then Err.Raise vbObjectError + 2500, , _
        "SOLIDWORKS is not running in this session. Nothing was changed."
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set source = swApp.ActiveDoc
    If source Is Nothing Then Err.Raise vbObjectError + 2501, , "No active SOLIDWORKS document."
    sourcePath = source.GetPathName
    If Len(sourcePath) = 0 Or fs.FileExists(sourcePath) = False Then _
        Err.Raise vbObjectError + 2502, , "The active document must be a saved assembly before kitify."
    sourceSize = fs.GetFile(sourcePath).Size
    sourceTime = fs.GetFile(sourcePath).DateLastModified
    RemoveMessageFilter swApp
    CloseUnsavedProbeDocuments swApp, sourcePath
    StartDecliner swApp, fs, "C:\skycode\logs\kitifyandboundingbox-declined.txt"

    stage = "Running kitify to completion before any bounding box"
    combined = KitifyDrifter("", "", False)
    If Len(combined) = 0 Or fs.FileExists(combined) = False Then _
        Err.Raise vbObjectError + 2503, , "Kitify did not return a saved combined assembly."
    If IsInsideFolder(fs, combined, fs.GetParentFolderName(sourcePath)) = False _
        Or StrComp(combined, sourcePath, vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 2504, , "Kitify wrote outside the source folder: " & combined
    AssertSourceUnchanged fs, sourcePath, sourceSize, sourceTime, "after kitify"
    RunKitifyAndBoundingBox = FinishWithBoundingBox(swApp, fs, sourcePath, sourceSize, sourceTime, combined, stage)
    StopDecliner
    RestoreMessageFilter
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = stage & ": " & Err.Description
    Debug.Print failureDescription
    WriteFailure failureDescription
    On Error Resume Next
    RestoreSavePrompts swApp, savedToggles
    StopDecliner
    RestoreMessageFilter
    On Error GoTo 0
    Err.Raise failureNumber, "KitifyAndBoundingBox", failureDescription
End Function

Private Function FinishWithBoundingBox(ByVal swApp As SldWorks.SldWorks, ByVal fs As Object, _
    ByVal sourcePath As String, ByVal sourceSize As Double, ByVal sourceTime As Date, _
    ByVal combined As String, ByRef stage As String) As String
    Dim generated As SldWorks.ModelDoc2
    Dim report As String
    Dim errors As Long
    Dim warnings As Long
    Dim savedToggles As Object
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed

    stage = "Reopening the saved combined assembly"
    Set savedToggles = Dictionary()
    SuppressSavePrompts swApp, savedToggles
    Set generated = swApp.OpenDoc6(combined, ASSEMBLY_DOCUMENT, 1, "", errors, warnings)
    If generated Is Nothing Or errors <> 0 Then Err.Raise vbObjectError + 2505, , _
        "Could not reopen combined output; errors=" & errors & "; warnings=" & warnings
    errors = 0
    If swApp.ActivateDoc3(generated.GetTitle, False, 0, errors) Is Nothing Then _
        Err.Raise vbObjectError + 2506, , "Could not activate combined output; errors=" & errors
    If StrComp(swApp.ActiveDoc.GetPathName, combined, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 2507, , "Refusing to add a bounding box to " & swApp.ActiveDoc.GetPathName
    ShowConfiguration swApp.ActiveDoc, PRODUCTION_CONFIGURATION

    stage = "Adding the bounding box to Default after kit configurations exist"
    report = GetCurrentModelBoundingBox(False)
    If swApp.ActivateDoc3(generated.GetTitle, False, 0, errors) Is Nothing _
        Or StrComp(swApp.ActiveDoc.GetPathName, combined, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2508, , "Bounding box did not leave the combined assembly active."
    End If
    ShowConfiguration generated, PRODUCTION_CONFIGURATION
    ConfineBoxToDefault generated
    VerifyBoxOnlyResolvedInDefault generated

    stage = "Saving the combined assembly without saving the source"
    errors = 0
    warnings = 0
    If generated.Save3(1, errors, warnings) = False Or errors <> 0 Then _
        Err.Raise vbObjectError + 2509, , "Could not save the boxed combined assembly; errors=" & errors & _
        "; warnings=" & warnings
    AssertSourceUnchanged fs, sourcePath, sourceSize, sourceTime, "after bounding-box save"
    ShowConfiguration generated, PRODUCTION_CONFIGURATION
    swApp.ActivateDoc3 generated.GetTitle, False, 0, errors
    WriteStepReport fs, combined, report, sourcePath, sourceSize, sourceTime
    RestoreSavePrompts swApp, savedToggles
    FinishWithBoundingBox = combined
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    RestoreSavePrompts swApp, savedToggles
    On Error GoTo 0
    Err.Raise failureNumber, "KitifyAndBoundingBox", failureDescription
End Function

Private Function IsInsideFolder(ByVal fs As Object, ByVal path As String, ByVal folder As String) As Boolean
    Dim parent As String
    parent = fs.GetParentFolderName(path)
    If StrComp(parent, folder, vbTextCompare) = 0 Then
        IsInsideFolder = True
    ElseIf Len(parent) > Len(folder) Then
        IsInsideFolder = CBool(StrComp(Left$(parent, Len(folder) + 1), folder & "\", vbTextCompare) = 0)
    End If
End Function

Private Function SourceBesideOutput(ByVal fs As Object, ByVal combined As String) As String
    Dim folderName As String
    Dim base As String
    Dim stampAt As Long
    folderName = fs.GetFileName(fs.GetParentFolderName(combined))
    If StrComp(Left$(folderName, 19), "kitified-generated_", vbTextCompare) <> 0 Then Exit Function
    base = Mid$(folderName, 20)
    stampAt = InStrRev(base, "_")
    If stampAt < 2 Then Exit Function
    base = Left$(base, stampAt - 1)
    stampAt = InStrRev(base, "_")
    If stampAt < 2 Or Len(Mid$(base, stampAt + 1)) <> 8 Then Exit Function
    If IsNumeric(Mid$(base, stampAt + 1)) = False Then Exit Function
    base = Left$(base, stampAt - 1)
    SourceBesideOutput = fs.BuildPath(fs.GetParentFolderName(fs.GetParentFolderName(combined)), base & ".SLDASM")
End Function

Private Sub CloseUnsavedProbeDocuments(ByVal swApp As SldWorks.SldWorks, ByVal sourcePath As String)
    Dim doc As SldWorks.ModelDoc2
    Dim nextDoc As SldWorks.ModelDoc2
    Dim title As String
    Dim path As String
    Set doc = swApp.GetFirstDocument
    Do While Not doc Is Nothing
        title = doc.GetTitle
        path = doc.GetPathName
        Set nextDoc = doc.GetNext
        If StrComp(path, sourcePath, vbTextCompare) <> 0 Then
            If Len(path) = 0 Or InStr(1, path, "box-config-probe-assembly.SLDASM", vbTextCompare) > 0 Then
                swApp.CloseDoc title
            End If
        End If
        Set doc = nextDoc
    Loop
End Sub

Private Sub AssertSourceUnchanged(ByVal fs As Object, ByVal sourcePath As String, _
    ByVal sourceSize As Double, ByVal sourceTime As Date, ByVal whenText As String)
    Dim file As Object
    Set file = fs.GetFile(sourcePath)
    If file.Size <> sourceSize Or file.DateLastModified <> sourceTime Then _
        Err.Raise vbObjectError + 2510, , "Source file changed " & whenText & ": " & sourcePath & _
        " size " & sourceSize & "->" & file.Size & " time " & CStr(sourceTime) & "->" & CStr(file.DateLastModified)
End Sub

Private Sub ConfineBoxToDefault(ByVal model As SldWorks.ModelDoc2)
    Dim names As Variant
    Dim name As Variant
    Dim component As SldWorks.Component2
    Dim feature As SldWorks.Feature
    Dim status As Long
    names = model.GetConfigurationNames
    If IsArray(names) = False Then Err.Raise vbObjectError + 2511, , "Combined configuration list is unavailable."
    For Each name In names
        If StrComp(CStr(name), PRODUCTION_CONFIGURATION, vbTextCompare) <> 0 Then
            ShowConfiguration model, CStr(name)
            Set component = FindBoxComponent(model)
            If Not component Is Nothing Then
                If CBool(component.IsSuppressed <> False) = False Then
                    status = component.SetSuppression2(0)
                    If status <> 2 Then Err.Raise vbObjectError + 2512, , _
                        "Could not suppress the bounding-box envelope in " & name & "; status=" & status
                End If
            End If
            Set feature = model.FeatureByName(BOX_FEATURE_NAME)
            If Not feature Is Nothing Then
                If CBool(feature.IsSuppressed <> False) = False Then
                    If feature.SetSuppression2(0, 1, Nothing) = False Then _
                        Err.Raise vbObjectError + 2513, , "Could not suppress " & BOX_FEATURE_NAME & " in " & name
                End If
            End If
        End If
    Next name
    ShowConfiguration model, PRODUCTION_CONFIGURATION
End Sub

Private Sub VerifyBoxOnlyResolvedInDefault(ByVal model As SldWorks.ModelDoc2)
    Dim names As Variant
    Dim name As Variant
    Dim component As SldWorks.Component2
    Dim feature As SldWorks.Feature
    names = model.GetConfigurationNames
    For Each name In names
        ShowConfiguration model, CStr(name)
        Set component = FindBoxComponent(model)
        Set feature = model.FeatureByName(BOX_FEATURE_NAME)
        If StrComp(CStr(name), PRODUCTION_CONFIGURATION, vbTextCompare) = 0 Then
            If component Is Nothing Or CBool(component.IsSuppressed <> False) Then _
                Err.Raise vbObjectError + 2514, , "Default does not have a resolved bounding-box envelope."
            If feature Is Nothing Or CBool(feature.IsSuppressed <> False) Then _
                Err.Raise vbObjectError + 2515, , "Default does not have a resolved " & BOX_FEATURE_NAME & " feature."
            If CBool(component.IsEnvelope <> False) = False Then _
                Err.Raise vbObjectError + 2516, , "The Default bounding-box component is not an envelope."
        Else
            If Not component Is Nothing Then
                If CBool(component.IsSuppressed <> False) = False Then _
                    Err.Raise vbObjectError + 2517, , "Bounding-box envelope is resolved in kit configuration " & name
            End If
            If Not feature Is Nothing Then
                If CBool(feature.IsSuppressed <> False) = False Then _
                    Err.Raise vbObjectError + 2518, , BOX_FEATURE_NAME & " is resolved in kit configuration " & name
            End If
        End If
    Next name
    ShowConfiguration model, PRODUCTION_CONFIGURATION
End Sub

Private Function FindBoxComponent(ByVal model As SldWorks.ModelDoc2) As SldWorks.Component2
    Dim root As SldWorks.Component2
    Dim children As Variant
    Dim child As Variant
    Dim component As SldWorks.Component2
    Set root = model.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    If root Is Nothing Then Exit Function
    children = root.GetChildren
    If IsArray(children) = False Then Exit Function
    For Each child In children
        Set component = child
        If InStr(1, component.GetPathName, "_BoundingBox.SLDPRT", vbTextCompare) > 0 _
            Or InStr(1, component.Name2, "BoundingBox", vbTextCompare) > 0 Then
            Set FindBoxComponent = component
            Exit Function
        End If
    Next child
End Function

Private Sub ShowConfiguration(ByVal model As SldWorks.ModelDoc2, ByVal name As String)
    If StrComp(model.ConfigurationManager.ActiveConfiguration.Name, name, vbTextCompare) = 0 Then Exit Sub
    If model.ShowConfiguration2(name) = False Then Err.Raise vbObjectError + 2519, , _
        "Could not activate configuration " & name
End Sub

Private Sub WriteFailure(ByVal description As String)
    Dim fs As Object
    Dim text As Object
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set text = fs.CreateTextFile("C:\skycode\logs\kitifyandboundingbox.txt", True, True)
    text.WriteLine "FAILED: " & description
    text.Close
End Sub

Private Sub WriteStepReport(ByVal fs As Object, ByVal combined As String, ByVal report As String, _
    ByVal sourcePath As String, ByVal sourceSize As Double, ByVal sourceTime As Date)
    Dim path As String
    Dim text As Object
    Dim file As Object
    path = fs.BuildPath(fs.GetParentFolderName(combined), "kitifyandboundingbox.txt")
    Set text = fs.CreateTextFile(path, True, True)
    text.WriteLine "Combined assembly: " & combined
    text.WriteLine "Active configuration at finish: Default"
    text.WriteLine "Source unchanged: " & sourcePath
    Set file = fs.GetFile(sourcePath)
    text.WriteLine "Source size: " & CStr(file.Size) & " (expected " & CStr(sourceSize) & ")"
    text.WriteLine "Source modified: " & CStr(file.DateLastModified) & " (expected " & CStr(sourceTime) & ")"
    text.WriteLine "Bounding box was added after kitify. Non-Default configurations have the envelope and feature suppressed."
    text.WriteLine "Deletion was not used: deleting the envelope from a kit configuration also removes it from Default."
    text.WriteLine "The 2026-11-19 Composer import did not see this suppressed box in kit configurations."
    text.WriteLine report
    text.Close
End Sub

Private Sub SuppressSavePrompts(ByVal swApp As SldWorks.SldWorks, ByVal saved As Object)
    SaveToggle swApp, saved, SW_TOGGLE_EXTREF_NO_PROMPT_OR_SAVE, True
    SaveToggle swApp, saved, SW_TOGGLE_WARN_SAVE_UPDATE_ERRORS, False
    SaveToggle swApp, saved, SW_TOGGLE_WARN_SAVING_REFERENCED_DOC, False
    SaveToggle swApp, saved, SW_TOGGLE_AUTO_DISMISS_OPEN_MESSAGES, True
End Sub

Private Sub SaveToggle(ByVal swApp As SldWorks.SldWorks, ByVal saved As Object, _
    ByVal preference As Long, ByVal value As Boolean)
    Dim current As Boolean
    current = CBool(swApp.GetUserPreferenceToggle(preference) <> False)
    If current = value Then Exit Sub
    saved.Add preference, current
    swApp.SetUserPreferenceToggle preference, value
End Sub

Private Sub RestoreSavePrompts(ByVal swApp As SldWorks.SldWorks, ByVal saved As Object)
    Dim key As Variant
    On Error Resume Next
    If saved Is Nothing Or swApp Is Nothing Then Exit Sub
    For Each key In saved.Keys
        swApp.SetUserPreferenceToggle CLng(key), CBool(saved(key))
    Next key
End Sub

Private Function Dictionary() As Object
    Set Dictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function AttachSolidWorks() As SldWorks.SldWorks
    On Error Resume Next
    Set AttachSolidWorks = GetObject(, "SldWorks.Application.33")
    If AttachSolidWorks Is Nothing Then Set AttachSolidWorks = GetObject(, "SldWorks.Application.34")
    On Error GoTo 0
End Function

Private Sub RemoveMessageFilter(ByVal swApp As SldWorks.SldWorks)
    If mFilterRemoved Then Exit Sub
    If swApp.GetProcessID = GetCurrentProcessId() Then Exit Sub
    If CoRegisterMessageFilter(0, mPreviousMessageFilter) = 0 Then mFilterRemoved = True
End Sub

Private Sub RestoreMessageFilter()
#If VBA7 Then
    Dim ignored As LongPtr
#Else
    Dim ignored As Long
#End If
    On Error Resume Next
    If mFilterRemoved Then CoRegisterMessageFilter mPreviousMessageFilter, ignored
    mPreviousMessageFilter = 0
    mFilterRemoved = False
End Sub

Private Sub StartDecliner(ByVal swApp As SldWorks.SldWorks, ByVal fs As Object, ByVal logPath As String)
    Dim script As Object
    Dim command As String
    If mDeclinerPid <> 0 Then Exit Sub
    mDeclinerScript = fs.BuildPath(fs.GetSpecialFolder(2), "skycode-kitifyandboundingbox-decline.ps1")
    Set script = fs.CreateTextFile(mDeclinerScript, True, True)
    script.Write DeclinerScript()
    script.Close
    command = "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -Command ""& ([scriptblock]::Create(" & _
        "[IO.File]::ReadAllText('" & Replace(mDeclinerScript, "'", "''") & "'))) " & swApp.GetProcessID & " " & _
        GetCurrentProcessId() & " '" & Replace(logPath, "'", "''") & "'"""
    mDeclinerPid = Shell(command, vbHide)
End Sub

Private Sub StopDecliner()
    Dim fs As Object
    On Error Resume Next
    If mDeclinerPid <> 0 Then
        GetObject("winmgmts:\\.\root\cimv2").Get("Win32_Process.Handle='" & CStr(mDeclinerPid) & "'").Terminate
        mDeclinerPid = 0
    End If
    If Len(mDeclinerScript) > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(mDeclinerScript) Then fs.DeleteFile mDeclinerScript, True
        mDeclinerScript = ""
    End If
End Sub

Private Function DeclinerScript() As String
    Dim s As String
    s = s & "param([int]$SwPid, [int]$HostPid, [string]$LogPath)" & vbCrLf
    s = s & "Add-Type -TypeDefinition @'" & vbCrLf
    s = s & "using System; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Text;" & vbCrLf
    s = s & "public static class SkycodeCheckoutDecliner {" & vbCrLf
    s = s & "  delegate bool EnumProc(IntPtr h, IntPtr l);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"")] static extern bool EnumWindows(EnumProc f, IntPtr l);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"")] static extern bool EnumChildWindows(IntPtr p, EnumProc f, IntPtr l);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"")] static extern bool IsWindowVisible(IntPtr h);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"")] static extern int GetDlgCtrlID(IntPtr h);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);" & vbCrLf
    s = s & "  [DllImport(""user32.dll"", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);" & vbCrLf
    s = s & "  public static bool IsTargetDialog(uint ownerPid, uint swPid, bool visible, string kind, string title) {" & vbCrLf
    s = s & "    return ownerPid == swPid && visible && kind == ""#32770"" && title.Equals(""SOLIDWORKS PDM"", StringComparison.OrdinalIgnoreCase);" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  public static bool IsNoButton(string kind, string text) {" & vbCrLf
    s = s & "    return kind.Equals(""Button"", StringComparison.OrdinalIgnoreCase) && text.Replace(""&"", """").Trim().Equals(""No"", StringComparison.OrdinalIgnoreCase);" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  public static List<string> Decline(uint pid) {" & vbCrLf
    s = s & "    var declined = new List<string>();" & vbCrLf
    s = s & "    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); var c = new StringBuilder(64); GetClassName(h, c, 64); var title = new StringBuilder(256); GetWindowText(h, title, 256);" & vbCrLf
    s = s & "      if (!IsTargetDialog(p, pid, IsWindowVisible(h), c.ToString(), title.ToString())) return true;" & vbCrLf
    s = s & "      var text = new StringBuilder(); IntPtr no = IntPtr.Zero; int noCount = 0;" & vbCrLf
    s = s & "      EnumChildWindows(h, (ch, l2) => { var t = new StringBuilder(2048); GetWindowText(ch, t, 2048); var kind = new StringBuilder(64); GetClassName(ch, kind, 64); if (IsNoButton(kind.ToString(), t.ToString())) { no = ch; noCount++; } text.Append(t).Append(' '); return true; }, IntPtr.Zero);" & vbCrLf
    s = s & "      if (noCount == 1) { string prefix = PostMessage(h, 0x0111, (IntPtr)(GetDlgCtrlID(no) & 0xFFFF), no) ? ""QUEUED NO: "" : ""ERROR: Could not queue No: ""; declined.Add(prefix + title + "": "" + text.ToString().Replace(""\r"", "" "").Replace(""\n"", "" "").Trim()); }" & vbCrLf
    s = s & "      return true; }, IntPtr.Zero); return declined; } }" & vbCrLf
    s = s & "'@" & vbCrLf
    s = s & "while ((Get-Process -Id $SwPid -ErrorAction SilentlyContinue) -and (Get-Process -Id $HostPid -ErrorAction SilentlyContinue)) {" & vbCrLf
    s = s & "  $declined = [SkycodeCheckoutDecliner]::Decline([uint32]$SwPid)" & vbCrLf
    s = s & "  foreach ($d in $declined) { Add-Content -LiteralPath $LogPath -Value ((Get-Date -Format s) + ' ' + $d) }" & vbCrLf
    s = s & "  if ($declined -match '^ERROR:') { break }" & vbCrLf
    s = s & "  if ($declined.Count -gt 0) { Start-Sleep -Milliseconds 1000 } else { Start-Sleep -Milliseconds 250 }" & vbCrLf
    s = s & "}" & vbCrLf
    DeclinerScript = s
End Function
