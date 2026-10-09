Attribute VB_Name = "VBAExtensibilityReferance"
Option Explicit

Sub AddVBAExtensibility()
    Dim ref As Object
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 1, 0)
    On Error GoTo 0
    If ref Is Nothing Then
        MsgBox "Microsoft VBA Extensibility is already added or could not be added."
    Else
        MsgBox "Microsoft VBA Extensibility added successfully."
    End If
End Sub

