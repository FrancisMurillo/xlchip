Attribute VB_Name = "TestChipInit"
Public Sub TestInstallChip()
On Error GoTo Cleanup
    ' Testing install
    ' 1. Create a fake workbook
    ' 2. Copy ChipInit via Import
    ' 3. Add references by hand, using this host machine's dlls
    ' 4. Run InstallChip on the other Workbook
    
    ' Setup
    Dim References As Variant ' Get references first before context/book switching
    References = ChipInit.ListProjectReferenceObjects
    
    Dim CurBook As Workbook, NewBook As Workbook
    Set CurBook = ActiveWorkbook
    Set NewBook = Workbooks.Add

    Dim Path As String
    Path = "~" & Format(Now(), "yyyymmddhhmmss") & "bas"
    
    CurBook.VBProject.VBComponents("ChipInit").Export Path
    NewBook.VBProject.VBComponents.Import Path
    DeleteFile Path
    
    Dim Ref As Variant
    For Each Ref In References
        If Not ChipInit.HasReference(Ref.Name, NewBook) Then
            NewBook.VBProject.References.AddFromFile Ref.FullPath
        End If
    Next
    
    Dim SampleBookPath As String
    SampleBookPath = CurBook.Path & Application.PathSeparator & "chip-TEST.xlsm"
    Application.Run NewBook.Name & "!ChipInit.InstallChip", SampleBookPath, False, False
    
    Dim ExpectedModules As Variant, ModuleName As Variant
    ExpectedModules = Array("Chip", "ChipInit", "ChipList")
    For Each ModuleName In ExpectedModules
        Debug.Print HasModule(CStr(ModuleName), NewBook)
    Next
    
Cleanup:
    If Not NewBook Is Nothing Then
        DoEvents
        NewBook.Close SaveChanges:=False
    End If
    
    If Err.Number <> 0 Then
        Stop
    End If
End Sub

Public Sub TestCheckDependency()
On Error Resume Next
    Debug.Print ChipInit.CheckDependencies
End Sub

Public Sub TestBrowseFile()
On Error Resume Next
    ' Do nothing, too intrusive to test
End Sub

Public Sub TestDownloadAndDeleteFile()
On Error Resume Next:
    Dim FilePath As String
    FilePath = ChipInit.DownloadFile() 'Download without a real path
    Debug.Print FilePath <> ""
    ChipInit.DeleteFile FilePath ' Delete temporary file
ErrHandler:
End Sub

Public Sub TestListProjectReferences()
On Error GoTo ErrHandler:
    Dim References As Variant
    References = ChipInit.ListProjectReferences
            
    Debug.Print UBound(References) + 1 = 7
ErrHandler:
End Sub
