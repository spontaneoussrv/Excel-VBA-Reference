Attribute VB_Name = "SeleniumBasicsReferance"
Sub AddSeleniumReference()
    Dim vbProj As Object
    Dim chkRef As Object
    Dim seleniumRefAdded As Boolean
    seleniumRefAdded = False
    Dim appDataPath As String
    
    ' Set the VBA project
    Set vbProj = ThisWorkbook.VBProject
    
    ' Check if Selenium Type Library is already added
    For Each chkRef In vbProj.References
        If chkRef.Name = "Selenium Type Library" Then
            seleniumRefAdded = True
            Exit For
        End If
    Next chkRef
    
    ' Add Selenium reference if not already added
    If Not seleniumRefAdded Then
        On Error Resume Next
        
        ' Try adding from the primary default path
        vbProj.References.AddFromFile "C:\Program files\seleniumBasic\Selenium32.tlb"
        
        ' If it fails, try adding from the AppData folder
        If Err.Number <> 0 Then
            ' Get the user's AppData folder path
            appDataPath = Environ("AppData") & "\seleniumBasic\Selenium32.tlb"
            vbProj.References.AddFromFile appDataPath
        End If
        
        On Error GoTo 0
        
        ' Check if the reference was added successfully
        If Err.Number = 0 Then
            MsgBox "Selenium Type Library reference added."
        Else
            MsgBox "Failed to add Selenium Type Library reference. Please check the installation paths."
        End If
    Else
        MsgBox "Selenium Type Library reference already exists."
    End If
End Sub

