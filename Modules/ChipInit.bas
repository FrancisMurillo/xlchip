Attribute VB_Name = "ChipInit"
'===========================
'Configurations
'===========================
'# This HTTP URL is where the Chip Workbook is stored
Private Const REPO_URL As String = "https://github.com/FrancisMurillo/chip/raw/0.1-poc/chip-RELEASE.xlsm"
Private Const DEPENDENCY_LIST As String = "Microsoft Visual Basic for Applications Extensibility *;Microsoft Scripting Runtime"
Private Const LIST_DELIMITER As String = ";"


'===========================
'Main Functions
'===========================

'# Install Chip by downloading the stable release file from the repository
'# and copying the required modules
Public Sub InstallChipFromRepo()
On Error GoTo ErrHandler
    Debug.Print "Install Chip From Repository"
    Debug.Print "=============================="
    
    Dim Path As String
    Debug.Print "Downloading Chip from " & REPO_URL
    Path = DownloadFile ' Download file using the default settings
    
    Debug.Print "Installing Chip"
    InstallChip Path ' Install Chip
    
    
    Debug.Print "Installation success"
Cleanup:
    If Path <> "" Then
        Debug.Print "Removing temporary file " & Path
        DeleteFile Path
    End If
    Exit Sub
ErrHandler:
    Debug.Print _
        "Whoops! There was an error in loading the file. " & _
        "Make sure you selected the URL is correctly pointed to a Chip workbook."
    Resume Cleanup
End Sub

Public Sub InstallChipLocally()
On Error GoTo ErrHandler
    Debug.Print "Install Chip Locally"
    Debug.Print "=============================="

    Dim Path As String
    Debug.Print "Select a Chip workbook"
    Path = BrowseFile
    If Path = "False" Then
        Debug.Print "No file was selected. Cancel installation"
        Exit Sub ' None was selected
    End If
    Debug.Print "Path: " & Path
    
    Debug.Print "Installing Chip"
    InstallChip Path ' Install Chip
    
    Debug.Print "Installation success"
    Exit Sub
ErrHandler:
    Debug.Print _
        "Whoops! There was an error in loading the file. " & _
        "Make sure you selected a Chip workbook."
End Sub

'===========================
'Internal Functions
'===========================

'# This copies the modules from the Chip workbook
'# The last core function
'@ Exception: Propagate
Private Sub InstallChip(ChipBookPath As String)
    Dim Dependencies As Variant
    Dependencies = Split(DEPENDENCY_LIST, ";")
    Debug.Print "Checking dependencies"
    If Not CheckDependencies(Dependencies) Then
        Debug.Print "One or more of the depedencies are not included. Make sure they are and installing again."
        Debug.Print "Required References:"
        For Each Depedency In Dependencies
            Debug.Print "# " & Depedency
        Next
        Err.Raise 1001
    End If
End Sub

'# This checks if the VB Project has the required references to run the code
'@ Param: Dependencies > A zero string array of dependencies
Public Function CheckDependencies(Dependencies As Variant) As Boolean
On Error GoTo ErrHandler
    Dim References As Variant
    References = ListProjectReferences
        
    Dim Depedency As Variant, Reference As Variant, IsFound As Boolean
    For Each Dependency In Dependencies
        IsFound = False
        For Each Reference In References
            IsFound = Reference Like Dependency
            If IsFound Then Exit For
        Next
        If Not IsFound Then
            CheckDependencies = False
            Exit Function
        End If
    Next
    CheckDependencies = True
ErrHandler:
End Function


'===========================
'Helper Functions
'===========================

'# This browses a file using the Open File Dialog
'# Primarily used to open a macro enabled file
'@ Exception: Propagated
'@ Return: The absolute path of the selected file, an "False" if none was selected
Public Function BrowseFile() As String
    BrowseFile = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
        FileFilter:="Excel Macro Enabled Files *.xlsm (*.xlsm),")
End Function

'# This downloads a file from the internet using the HTTP GET method
'# This is primarily used for downloading a binary file or the workbook repo needed
'! Taken from a site, modified to my use
'@ Exception: Propagated
'@ Return: The absolute path of the downloaded file, if path was not provided else the path itself
Public Function DownloadFile(Optional URL As String = REPO_URL, Optional Path As String = "")
On Error GoTo ErrHandler:
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
ErrHandler:
End Function

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
