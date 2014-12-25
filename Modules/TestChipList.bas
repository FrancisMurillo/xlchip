Attribute VB_Name = "TestChipList"
Public Sub TestChipNames()
On Error Resume Next
    Dim ExpectedChips As Variant, ActualChips As Variant
    ExpectedChips = Array("Vase", "Wheat")
    ActualChips = ChipList.ListChipNames()
    
    VaseAssert.AssertEqualArrays ActualChips, ExpectedChips
End Sub
