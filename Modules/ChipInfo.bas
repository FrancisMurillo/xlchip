Attribute VB_Name = "ChipInfo"
Private Reference As Variant
Private Modules As Variant

Public Sub Initialize()
    Reference = Array( _
        "Microsoft Visual Basic for Applications Extensibility *", _
        "Microsoft Scripting Runtime")
    Modules = Array( _
        "Chip", "ChipLib", "ChipList")
End Sub
