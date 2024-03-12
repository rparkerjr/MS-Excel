Attribute VB_Name = "FreeFlow"
'FreeFlow2.0 Module
'
'Author:        Richard Parker
'Version:       0.9b        2024-02-13

Sub RefreshAllData()

    ThisWorkbook.RefreshAll

End Sub

Sub InitNewProjectForm()
    Dim destructList As Variant
    destructList = Range("DestructStatus").Value
    Load NewProject
    NewProject.Show
    'Unload NewProject
End Sub

Sub InitTabListing()
    Call PopulateTabListing
    Load TabListing
    TabListing.Show
End Sub

Public Sub PopulateTabListing()
    Dim numSheets As Integer
    numSheets = ActiveWorkbook.Worksheets.Count - 1
    Dim sheetList() As String
    ReDim sheetList(numSheets, 1)
    
    For i = 0 To numSheets
        sheetList(i, 0) = ActiveWorkbook.Worksheets(i + 1).Name
        If ActiveWorkbook.Worksheets(i + 1).Visible Then
            sheetList(i, 1) = "Visible"
        Else
            sheetList(i, 1) = "Hidden"
        End If
    Next i
    
    TabListing.tabListBox.Clear
    TabListing.tabListBox.ColumnCount = 2
    TabListing.tabListBox.list = sheetList
End Sub


