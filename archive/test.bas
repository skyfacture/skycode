Attribute VB_Name = "ComposerConnectionTest"
Option Explicit

Public Sub TestComposer()
    Dim swApp As Object
    Dim swDoc As Object

    Set swApp = TryCreateComposerApp("DSComposer.Application")
    If swApp Is Nothing Then
        Set swApp = TryCreateComposerApp("Composer.Application")
    End If

    If swApp Is Nothing Then
        MsgBox "Could not create Composer COM application using DSComposer.Application or Composer.Application."
        Exit Sub
    End If

    MsgBox "Connected to Composer!"

    Set swDoc = swApp.GetActiveDocument
    If Not swDoc Is Nothing Then
        MsgBox "Active document has " & swDoc.ViewCount & " views."
    Else
        MsgBox "No active document."
    End If
End Sub

Private Function TryCreateComposerApp(ByVal progId As String) As Object
    On Error Resume Next
    Set TryCreateComposerApp = CreateObject(progId)
    On Error GoTo 0
End Function