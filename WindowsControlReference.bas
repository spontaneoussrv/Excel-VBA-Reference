Attribute VB_Name = "WindowsControlReferance"
Option Explicit

Sub AddWindowsCommonControls()
    Dim ref As Object
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References.AddFromFile("C:\Windows\System32\MSCOMCTL.OCX")
    On Error GoTo 0
    If ref Is Nothing Then
        MsgBox "Microsoft Windows Common Controls is already added or could not be added."
    Else
        MsgBox "Microsoft Windows Common Controls added successfully."
    End If
End Sub

