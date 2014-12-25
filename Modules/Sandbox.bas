Attribute VB_Name = "Sandbox"
Public gX As Variant

Public Sub TryPath()
    ChipLib.DeleteModule "ChipInfo", ActiveWorkbook
    ChipLib.DeleteModule "ImpMod1", ActiveWorkbook
    ChipLib.DeleteModule "ImpModB", ActiveWorkbook

    ChipLib.AttachChip "C:\Users\NOBODY\Desktop\Robot\Workspace\Excel\Chip\sample-chip.xlsm"
    
    

    ChipLib.DeleteModule "ChipInfo", ActiveWorkbook
    ChipLib.DeleteModule "ImpMod1", ActiveWorkbook
    ChipLib.DeleteModule "ImpModB", ActiveWorkbook
End Sub
Public Sub DynamicInvoke()
    Dim Path As String
    Path = ActiveWorkbook.Path & Application.PathSeparator & "ChipInfo.bas"
    ActiveWorkbook.VBProject.VBComponents.Import Path

    ChipReadInfo.ClearInfo
    Application.Run "ChipInfo.WriteInfo"
    DoEvents

    ActiveWorkbook.VBProject.VBComponents.Remove _
        ActiveWorkbook.VBProject.VBComponents("ChipInfo")
End Sub

Public Sub X()
    Dim X
    'x = ChipInit.ListProjectReferences
    'x = ChipInit.BrowseFile
    Dim Dependencies As Variant
    Dependencies = Array( _
        "Microsoft Visual Basic for Applications Extensibility *" _
      , "Microsoft Scripting Runtime" _
        )
    'x = ChipList.RepositoryList
    'x = ChipInit.CheckDependencies(Dependencies)
    X = ChipInit.ListWorkbookModules(ActiveWorkbook)
    ActiveWorkbook.VBProject.VBComponents ("Chip")
End Sub

Sub GetModules()
Dim modName As String
Dim wb As Workbook
Dim l As Long

Set wb = ThisWorkbook

For l = 1 To wb.VBProject.VBComponents.Count
    With wb.VBProject.VBComponents(l)
        modName = modName & vbCr & .Name
    End With
Next

MsgBox "Module Names:" & vbCr & modName

Set wb = Nothing

End Sub
