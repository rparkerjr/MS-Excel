Attribute VB_Name = "FreeFlow"
'FreeFlow2.0 Module
'
'Author:        Richard Parker
'Version:       0.9b        2024-02-13
'               1.0         2024-03-11
'               1.1         2024-03-12
'               1.2         2024-03-13
'               1.3         2024-03-23

Sub RefreshAllData()

    ThisWorkbook.RefreshAll

End Sub

Sub InitNewProjectForm()
    Dim destructList As Variant
    destructList = Range("DestructStatus").Value
    Load NewProject
    NewProject.Show
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

Public Sub AddBoxesToProject(newProjectSheet As Worksheet, startNum As Long, endNum As Long)
    Dim tbl As ListObject
    Dim tblRow As ListRow
    
    Set tbl = newProjectSheet.ListObjects(1)
    
    For box = startNum To endNum
        Set tblRow = tbl.ListRows.Add
        tblRow.Range(1, 2).Value = box
        tblRow.Range(1, 3).Value = "Received"
    Next
    
    newProjectSheet.PivotTables("SUMMARY_PIVOT").RefreshTable
    
    Set tbl = Nothing
    Set tblRow = Nothing
End Sub

Public Sub AddRowToMaster(tabName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim newRowProject As Range
    
    Set ws = Worksheets("Master Tracking")
    Set tbl = ws.ListObjects("MASTER")
    Set tblRow = tbl.ListRows.Add
    tblRow.Range(1, 4).Value = tabName
    
    Set newRowProject = activeCell
    
    Set ws = Nothing
    Set tbl = Nothing
    Set tblRow = Nothing
End Sub
