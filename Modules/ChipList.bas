Attribute VB_Name = "ChipList"
Public RepositoryList As Variant

Public Sub ReloadRepositoryList()
    ' This is a list of tuples containing the following
    ' Chip Name, URL
    RepositoryList = Array( _
          Array("Vase", "https://github.com/FrancisMurillo/vase/raw/master/vase-RELEASE.xlsm") _
        , Array("Wheat", "https://github.com/FrancisMurillo/wheat/raw/master/wheat-RELEASE.xlsm") _
    )
End Sub



