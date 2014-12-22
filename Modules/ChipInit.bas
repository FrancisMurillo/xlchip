Attribute VB_Name = "ChipInit"
'===========================
'Configurations
'===========================

'# This HTTP URL is where the Chip Workbook is stored
Private Const REPO_URL As String = ""


'===========================
'Main Functions
'===========================

'# Install Chip by downloading the stable release file from the repository
'# and copying the required modules
Public Sub InstallChipFromRepo()

End Sub

Public Sub InstallChipLocally()

End Sub

'===========================
'Internal Functions
'===========================

Private Sub InstallChip()

End Sub

Private Sub CheckDependencies()

End Sub


'===========================
'Helper Functions
'===========================
Sub Test()
    Dim FileNum As Long
    Dim FileData() As Byte
    Dim MyFile As String
    Dim WHTTP As Object
    
    On Error Resume Next
        Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5")
        If Err.Number <> 0 Then
            Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
        End If
    On Error GoTo 0
    
    MyFile = "http://www.vba-and-excel.com/vba/internet/6-loading-information-from-the-internet-using-the-xmlhttp-object"
    MyFile = "http://raw.githubusercontent.com/FrancisMurillo/chip/master/LICENSE"
    MyFile = "http://p2p.wrox.com/image.php?u=402&dateline=1229009937"
    
    WHTTP.Open "GET", MyFile, False
    WHTTP.Send
    FileData = WHTTP.responseBody
    Set WHTTP = Nothing
    
    If Dir("C:\MyDownloads", vbDirectory) = Empty Then MkDir "C:\MyDownloads"
    
    FileNum = FreeFile
    Open "C:\Users\NOBODY\Desktop\Robot\a.jpg" For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
    
    MsgBox "Open the folder [ C:\MyDownloads ] for the downloaded file..."
End Sub

'# Deletes a file forcibly, it does not check whether it is a folder or the path does not exists
'# This is used to delete a temp file whether it still exists or not
'@ Exception: Propagated
Public Sub DeleteFile(FilePath As String)
On Error Resume Next
    With New FileSystemObject
        If .FileExists(FilePath) Then
            .DeleteFile FilePath
        End If
    End With
End Sub

'# This returns an string array of the references used in this VBA Project
'# The strings are the name of the references, not the filename or path
'@ Return: A zero-based array of strings
'@ Exception: Propagated
Public Function ListProjectReferences() As Variant
On Error Resume Next
    Dim VBProj As VBIDE.VBProject
    Set VBProj = Application.VBE.ActiveVBProject
    
    Dim ReferenceLength As Integer, Index As Long
    Dim References As Variant
    
    ReferenceLength = VBProj.References.Count
    If ReferenceLength = 0 Then
        ListProjectReferences = Array()
        Exit Function
    End If
    
    References = Array()
    ReDim References(1 To ReferenceLength)
    For Index = 1 To VBProj.References.Count
        With VBProj.References.Item(Index)
            References(Index) = .Description
        End With
    Next
    
    ReDim Preserve References(0 To ReferenceLength - 1)
    ListProjectReferences = References
End Function
