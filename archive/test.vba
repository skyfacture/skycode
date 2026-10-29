Sub TestComposer()
    Dim swApp As Object
    Dim swDoc As Object
    Set swApp = CreateObject("DSComposer.Application")
    MsgBox "Connected to Composer!"
    Set swDoc = swApp.GetActiveDocument
    If Not swDoc Is Nothing Then
        MsgBox "Active document has " & swDoc.ViewCount & " views."
    Else
        MsgBox "No active document."
    End If
End Sub