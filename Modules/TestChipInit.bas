Attribute VB_Name = "TestChipInit"
Public Sub TestCheckDependency()
On Error Resume Next
    Debug.Print ChipInit.CheckDependencies
End Sub

Public Sub TestBrowseFile()
On Error Resume Next
    ' Do nothing, too intrusive to test
End Sub

Public Sub TestDownloadAndDeleteFile()
On Error GoTo ErrHandler:
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
