VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewProject 
   Caption         =   "Create New Production Project"
   ClientHeight    =   7140
   ClientLeft      =   75
   ClientTop       =   270
   ClientWidth     =   11100
   OleObjectBlob   =   "NewProject.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "NewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelButton_Click()
NewProject.Hide
Unload NewProject
End Sub

Private Sub createProjectButton_Click()
Dim template As Worksheet
Dim newSheet As Worksheet
Dim newList As ListObject
Dim newPivot As PivotTable
On Error Resume Next

'Validate work order before starting
'workOrder.Value

Application.ScreenUpdating = False

Set template = ActiveWorkbook.Sheets("ClientProject")
template.Copy Before:=Worksheets("ClientProject")
Set newSheet = Application.ActiveSheet

'Convert spaces to underscores
Dim cleanedTabName As String
cleanedTabName = Replace(tabName.Value, " ", "_")

With newSheet
    .Name = cleanedTabName
    .Range("Work_Order").Value = workOrder.Value
    .Range("Client_Name").Value = clientName.Value
    .Range("Department_Name").Value = department.Value
    .Range("Client_Project").Value = cleanedTabName
    .Range("Project_Status").Value = "Received"
    .Range("Shred").Value = shredCombo.Value
    .Range("Contact_Name").Value = contact.Value
    .Range("Date_Received").Value = pickupDate.Value
    .Range("Pickup_By").Value = pickupBy.Value
    .Range("Last_Update").Value = Date
    .Range("Updated_By").Value = Application.UserName
    .Range("Notes").Value = otherNotes.Value
End With

'Relink new pivot table and data source
Set newPivot = newSheet.PivotTables(1)
Set newList = newSheet.ListObjects(1)
newList.Name = "BOXES_" & UCase(cleanedTabName)
newPivot.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newList, Version:=xlPivotTableVersion15)

'Need to generate box numbers automatically?

'Dim newRow As ListRow
'ActiveWorkbook.Worksheets("Master Tracking").ListObject("MASTER").ListRows.Add AlwaysInsert:=True
'newRow.Range.Value = tabName.Value

Application.ScreenUpdating = True

Unload NewProject

Set template = Nothing
Set newSheet = Nothing
Set newPivot = Nothing
Set newList = Nothing
Set newRow = Nothing
End Sub

