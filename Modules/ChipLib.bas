Attribute VB_Name = "ChipLib"

'# Clears the intermediate screen
Public Sub ClearScreen()
    Application.SendKeys "^g ^a {DEL}"
End Sub

'# Removes a module whether it exists or not
'# Used in making sure there are no duplicate modules
Public Sub DeleteModule(ModuleName As String, Book As Workbook)
On Error Resume Next
    Dim CurProj As VBProject, Module As VBComponent
    Set CurProj = Book.VBProject
    Set Module = CurProj.VBComponents(ModuleName)
    CurProj.VBComponents.Remove Module
    DoEvents
    Err.Clear
End Sub

'# Checks if an module exists
Public Function HasModule(ModuleName As String, Book As Workbook) As Boolean
On Error Resume Next
    HasModule = False
    HasModule = Not ActiveWorkbook.VBProject.VBComponents(ModuleName) Is Nothing  ' This fails if the module does not exists thus defaulting to False
    Err.Clear
End Function

'# Lists the modules of an workbook
'# Primarily used to get all Chip modules
'@ Return: An array of VB Components
Public Function ListWorkbookModuleObjects(Book As Workbook) As Variant
    Dim Comp As VBComponent, Modules As Variant, Index As Long
    Modules = Array()
    ReDim Modules(0 To Book.VBProject.VBComponents.Count - 1)
    For Each Comp In Book.VBProject.VBComponents
        Set Modules(Index) = Comp
        Index = Index + 1
    Next
    ListWorkbookModuleObjects = Modules
End Function

'# This browses a file using the Open File Dialog
'# Primarily used to open a macro enabled file
'@ Return: The absolute path of the selected file, an "False" if none was selected
Public Function BrowseFile() As String
    BrowseFile = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
        FileFilter:="Excel Macro Enabled Files *.xlsm (*.xlsm),")
End Function

'# This downloads a file from the internet using the HTTP GET method
'# This is primarily used for downloading a binary file or the workbook repo needed
'! Taken from a site, modified to my use
'@ Return: The absolute path of the downloaded file, if path was not provided else the path itself
Public Function DownloadFile(Optional URL As String = REPO_URL, Optional Path As String = "")
    If Path = "" Then ' Create pseudo unique path
        Path = ActiveWorkbook.Path & Application.PathSeparator & "~" & Format(Now(), "yyyymmddhhmmss")
    End If

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
    
    WHTTP.Open "GET", URL, False
    WHTTP.Send
    FileData = WHTTP.responseBody
    Set WHTTP = Nothing
    
    FileNum = FreeFile
    Open Path For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
    
    DownloadFile = Path
    Exit Function
End Function

'# Deletes a file forcibly, it does not check whether it is a folder or the path does not exists
'# This is used to delete a temp file whether it still exists or not
Public Sub DeleteFile(FilePath As String)
    With New FileSystemObject
        If .FileExists(FilePath) Then
            .DeleteFile FilePath
        End If
    End With
End Sub

'# This returns an string array of the references used in this VBA Project
'# The strings are the name of the references, not the filename or path
'@ Return: A zero-based array of strings
Public Function ListProjectReferences() As Variant
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

'# This returns an array of project references, the objects themselves for use
'# This is used for setting up the test workbook to have the correct references
'@ Return: A zero-based array of references
Public Function ListProjectReferenceObjects() As Variant
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
        Set References(Index) = VBProj.References.Item(Index)
    Next
    
    ReDim Preserve References(0 To ReferenceLength - 1)
    ListProjectReferenceObjects = References
End Function

'# Checks if the refrence exists for a workbook given its name
Public Function HasReference(ReferenceName As String, Book As Workbook) As Boolean
On Error Resume Next
    HasReference = False
    HasReference = Not Book.VBProject.References(ReferenceName) Is Nothing
    Err.Clear
End Function


