Attribute VB_Name = "HTML_XML_Referaance"
Sub Add_HTML_XML_References()
    Dim vbProj As Object
    Dim chkRef As Object
    
    ' Get the active VBA project
    Set vbProj = ThisWorkbook.VBProject
    
    ' Check and add "Microsoft HTML Object Library"
    On Error Resume Next
    Set chkRef = vbProj.References("Microsoft HTML Object Library")
    If chkRef Is Nothing Then
        vbProj.References.AddFromGuid "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}", 4, 0
    End If
    On Error GoTo 0
    
    ' Check and add "Microsoft XML, v6.0"
    On Error Resume Next
    ' Attempt to add the XML reference directly by GUID
    vbProj.References.AddFromGuid "{F5078F18-C551-11D3-89B9-0000F81FE221}", 6, 0
    On Error GoTo 0
    
    MsgBox "References checked and added if necessary.", vbInformation
End Sub

