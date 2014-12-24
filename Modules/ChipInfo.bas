Attribute VB_Name = "ChipInfo"
Public Sub WriteInfo()
    ChipReadInfo.References = Array( _
        "Microsoft Visual Basic for Applications Extensibility *")
    ChipReadInfo.Modules = Array( _
        "Vase", "VaseLib", "VaseAssert", "VaseConfig")
End Sub
