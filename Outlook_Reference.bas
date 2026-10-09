Attribute VB_Name = "Outlook_Referance"
Sub AddOutlookLibrary()
    Dim ref As Object
    On Error Resume Next
    Set ref = ThisWorkbook.VBProject.References.AddFromGuid("{00062FFF-0000-0000-C000-000000000046}", 1, 0)
    On Error GoTo 0
    If ref Is Nothing Then
        MsgBox "Microsoft Outlook Object Library is already added or could not be added."
    Else
        MsgBox "Microsoft Outlook Object Library added successfully."
    End If
End Sub

