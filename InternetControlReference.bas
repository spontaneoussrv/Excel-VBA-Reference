Attribute VB_Name = "InternetControlReferance"
Sub AddInternetControlsReference()
    Dim ref As Object
    Dim refName As String
    Dim refPath As String
    
    refName = "Microsoft Internet Controls"
    refPath = "C:\Windows\System32\ieframe.dll" ' Path may vary based on system and version
    
    ' Check if the reference is already added
    On Error Resume Next
    Set ref = Nothing
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Name = refName Then
            MsgBox "Reference already added."
            Exit Sub
        End If
    Next ref
    On Error GoTo 0
    
    ' Add the reference
    ThisWorkbook.VBProject.References.AddFromFile refPath
    MsgBox "Reference added successfully."
End Sub
