Attribute VB_Name = "TestChipInit"
Public Sub TestDeleteFile()
On Error GoTo ErrHandler:
    
ErrHandler:
End Sub

Public Sub TestListProjectReferences()
On Error GoTo ErrHandler:
    Dim References As Variant
    References = ChipInit.ListProjectReferences
            
    Debug.Print UBound(References) + 1 = 7
ErrHandler:
End Sub
