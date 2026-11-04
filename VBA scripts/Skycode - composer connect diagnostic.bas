Attribute VB_Name = "ComposerConnectDiagnostic"
Option Compare Database
Option Explicit

' ============================================================
' Skycode - Composer Connection Diagnostic
' Run DiagComposer() from the Access VBA immediate window / macro.
' Writes logs\composer diagnostic.txt and shows a summary.
'
' Purpose: isolate why the .bas scripts fail on Composer 2025
' but worked on 2024 (COM registration, ProgID, typelib reference,
' running-instance attach vs new instance).
' ============================================================

Private Const LOG_PATH As String = "C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\composer diagnostic.txt"

Public Sub DiagComposer()
    Dim fs As Object, logFile As Object
    Dim swApp As Object
    Dim swDoc As Object
    Dim summary As String

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set logFile = fs.CreateTextFile(LOG_PATH, True)

    Log logFile, "=== Composer connection diagnostic " & Now & " ==="
    Log logFile, "Host: " & Application.Name & " " & Application.Version & _
                 "  |  VBA7=" & CStr(IsVba7()) & "  |  Win64=" & CStr(IsWin64())

    ' 1) What is registered for each candidate ProgID
    DumpProgId logFile, "DSComposer.Application"
    DumpProgId logFile, "Composer.Application"
    DumpProgId logFile, "ComposerPlayer.Application"

    ' 2) Try to ATTACH to an already-running Composer (preferred)
    Set swApp = AttachComposer(logFile)

    ' 3) Fall back to launching a new instance
    If swApp Is Nothing Then Set swApp = LaunchComposer(logFile)

    If swApp Is Nothing Then
        summary = "No Composer COM object could be obtained. See " & LOG_PATH
        Log logFile, summary
        logFile.Close
        MsgBox summary, vbCritical
        Exit Sub
    End If

    ' 4) Probe the object
    On Error Resume Next
    Log logFile, "TypeName(swApp) = " & TypeName(swApp)
    If Err.Number <> 0 Then Log logFile, "TypeName failed: " & Err.Number & " " & Err.Description
    Err.Clear

    Log logFile, "swApp.Version = " & CStr(swApp.Version)
    If Err.Number <> 0 Then Log logFile, ".Version failed: " & Err.Number & " " & Err.Description
    Err.Clear

    Set swDoc = swApp.GetActiveDocument
    If Err.Number <> 0 Then
        Log logFile, "GetActiveDocument failed: " & Err.Number & " " & Err.Description
        Err.Clear
    End If

    If swDoc Is Nothing Then
        Log logFile, "GetActiveDocument returned Nothing (this is the classic cause of run-time error 91 " & _
                     "further down the scripts - a NEW instance has no open document)."
    Else
        Log logFile, "Active document OK. ViewCount = " & CStr(swDoc.ViewCount)
        If Err.Number <> 0 Then Log logFile, ".ViewCount failed: " & Err.Number & " " & Err.Description
        Err.Clear

        Log logFile, "GetAllViews raw = '" & CStr(swDoc.GetAllViews) & "'"
        If Err.Number <> 0 Then Log logFile, ".GetAllViews failed: " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    summary = "Diagnostic complete. See " & LOG_PATH
    Log logFile, summary
    logFile.Close
    MsgBox summary, vbInformation
End Sub

' --- Attach to a running Composer instance (ROT) -------------
Public Function AttachComposer(Optional ByVal logFile As Object) As Object
    Dim o As Object
    Dim ids As Variant, i As Long
    ids = Array("DSComposer.Application", "Composer.Application")

    For i = LBound(ids) To UBound(ids)
        On Error Resume Next
        Set o = GetObject(, CStr(ids(i)))
        If Err.Number = 0 And Not o Is Nothing Then
            Log logFile, "GetObject(, """ & ids(i) & """) -> attached to running instance."
            On Error GoTo 0
            Set AttachComposer = o
            Exit Function
        Else
            Log logFile, "GetObject(, """ & ids(i) & """) failed: " & Err.Number & " " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Function

' --- Launch a new Composer instance --------------------------
Private Function LaunchComposer(ByVal logFile As Object) As Object
    Dim o As Object
    Dim ids As Variant, i As Long
    ids = Array("DSComposer.Application", "Composer.Application")

    For i = LBound(ids) To UBound(ids)
        On Error Resume Next
        Set o = CreateObject(CStr(ids(i)))
        If Err.Number = 0 And Not o Is Nothing Then
            Log logFile, "CreateObject(""" & ids(i) & """) -> new instance created."
            On Error GoTo 0
            Set LaunchComposer = o
            Exit Function
        Else
            Log logFile, "CreateObject(""" & ids(i) & """) failed: " & Err.Number & " " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Function

' --- Registry inspection -------------------------------------
Private Sub DumpProgId(ByVal logFile As Object, ByVal progId As String)
    Dim wsh As Object
    Dim clsid As String
    Dim exePath As String

    Set wsh = CreateObject("WScript.Shell")

    On Error Resume Next
    clsid = wsh.RegRead("HKCR\" & progId & "\CLSID\")
    If Err.Number <> 0 Then
        Log logFile, "ProgID '" & progId & "' NOT registered."
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear

    exePath = wsh.RegRead("HKCR\CLSID\" & clsid & "\LocalServer32\")
    If Err.Number <> 0 Then
        exePath = wsh.RegRead("HKCR\CLSID\" & clsid & "\InprocServer32\")
    End If
    If Err.Number <> 0 Then
        exePath = "(no LocalServer32/InprocServer32 value - registration is incomplete)"
        Err.Clear
    End If
    On Error GoTo 0

    Log logFile, "ProgID '" & progId & "' -> CLSID " & clsid & " -> " & exePath
End Sub

' --- Helpers -------------------------------------------------
Private Sub Log(ByVal logFile As Object, ByVal msg As String)
    Debug.Print msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

Private Function IsVba7() As Boolean
#If VBA7 Then
    IsVba7 = True
#End If
End Function

Private Function IsWin64() As Boolean
#If Win64 Then
    IsWin64 = True
#End If
End Function
