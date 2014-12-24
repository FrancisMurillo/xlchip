Attribute VB_Name = "Chip"
'' These modules contain the functions required to install the other modules

'===========================
'Main Functions
'===========================

'# Install a Chip module using an name from the list or an direct URL
'# Using the custom parameters, URL takes precedence over name if both are provided
Public Sub ChipOnFromRepo(Optional ChipRepo As String = "", Optional URL As String = "")
On Error GoTo ErrHandler
    ChipLib.ClearScreen
    
    ' Predeclaration for errors
    Dim ChipBook As Workbook, CurBook As Workbook
    Dim ChipPath As String
    
    Debug.Print "Chip On From Repository"
    Debug.Print "=============================="
    
    If ChipRepo = "" And URL = "" Then
        Debug.Print "Enter a Chip Repo or an URL to attach a chip"
        Exit Sub
    End If
     
    If URL = "" Then ' Download file using the list
        Debug.Print "Finding a Chip using the list"
        Debug.Print "Looking for " & ChipRepo
        
        URL = ChipList.FindURLByName(ChipRepo)
        If URL = "" Then
            Debug.Print "Cannot find " & ChipRepo & " in the Chip Repository List"
            Debug.Print "Available Chips:"
            
            Dim ChipName As Variant
            For Each ChipName In ChipList.ListChipNames
                Debug.Print "* " & ChipName
            Next

            Debug.Print ""
            Debug.Print "Enter an available repo or update the list"
            Exit Sub
        End If
    Else
        Debug.Print "Finding a Chip using an URL"
    End If
    
    Debug.Print "Using URL " & URL
    Debug.Print "Downloading file, this might take a while"
    ChipPath = ChipLib.DownloadFile(URL)
     
    If ChipPath = "" Then
        Debug.Print "There was error downloading the file. Make sure the URL is valid before trying again."
        Exit Sub
    End If
    
    Debug.Print "File downloaded successfully"
    Debug.Print ""
    
    ChipLib.AttachChip ChipPath
    
    Debug.Print ""
    Debug.Print "Chip installed"
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Whoops! There was an error installing the module." & _
            "Double check if the book and chip is okay."
    End If
    Err.Clear

    Debug.Print "Cleaning up temporary file"
    ChipLib.DeleteFile ChipPath
End Sub

'# Install chip locally, which is unusual but the base use case if downloading does not work
Public Sub ChipOnLocally(Optional Path As String = "")
On Error GoTo ErrHandler
    ChipLib.ClearScreen
    
    Debug.Print "Chip On Locally"
    Debug.Print "=============================="
    
    If Path = "" Then
        Debug.Print "Browse the chip you want to install"
        Path = ChipLib.BrowseFile
        If Path = "" Then
            Debug.Print "Canceled browsing. Changed your mind?"
            Exit Sub
        End If
    Else
        Debug.Print "Chip path was given"
        If Not ChipLib.DoesFileExists(Path) Then
            Debug.Print "File path does not exist. Enter the path carefully or verify it is correct."
            Exit Sub
        End If
    End If
    
    Debug.Print "Using the path " & Path
    Debug.Print ""
    
    ChipLib.AttachChip Path
    
    Debug.Print ""
    Debug.Print "Chip installed"
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Whoops! There was an error installing the module." & _
            "Double check if the book and chip is okay."
    End If
    Err.Clear
End Sub

