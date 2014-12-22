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
Public Sub DownloadFile()
Dim myURL As String
myURL = "http://raw.githubusercontent.com/FrancisMurillo/chip/master/LICENSE"
myURL = "http://www.vba-and-excel.com/vba/internet/6-loading-information-from-the-internet-using-the-xmlhttp-object"

Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False
WinHttpReq.Send

myURL = WinHttpReq.responseBody
If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile "C:\License"
    oStream.Close
End If

End Sub

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
'# It simply makes sure the file on that path is deleted
Public Sub DeleteFile(FilePath As String)
On Error GoTo ErrHandler
    With New FileSystemObject
        If .FileExists(FilePath) Then
            .DeleteFile FilePath
        End If
    End With
    Exit Sub
ErrHandler:
    Err.Clear
End Sub
