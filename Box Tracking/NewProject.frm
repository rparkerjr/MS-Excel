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

Private Sub UserForm_Initialize()
   Dim maxValue
   maxValue = WorksheetFunction.Max(Range("Work_Order_Numbers"))
   If Not maxValue > 0 Then Exit Sub
   Me.workOrder.Value = maxValue + 1
End Sub

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

If ValidateWorkOrder(workOrder.Value) = False Then GoTo Return_to_Form
If ValidateBoxNumbers(boxRangeStart.Value, boxRangeEnd.Value) = False Then GoTo Return_to_Form

Application.ScreenUpdating = False

Set template = ActiveWorkbook.Sheets("ClientProject")
template.Copy Before:=Worksheets("ClientProject")
Set newSheet = Application.ActiveSheet

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

Call AddBoxesToProject(newSheet, CLng(boxRangeStart.Value), CLng(boxRangeEnd.Value))

AddRowToMaster (cleanedTabName)

Unload NewProject

Return_to_Form:
Application.ScreenUpdating = True
Set template = Nothing
Set newSheet = Nothing
Set newPivot = Nothing
Set newList = Nothing
Set newRow = Nothing
End Sub

Private Function ValidateWorkOrder(workOrderNumber As String) As Boolean
Dim workOrderList As Range

If Trim(workOrder.Value) = "" Then
    ValidateWorkOrder = False
    MsgBox ("Please enter a Work Order Number.")
    Exit Function
End If

'Set workOrderList = Sheets("Master Tracking").ListObjects("MASTER").ListColumns("Work Order Number").Range()
Set workOrderList = Range("Work_Order_Numbers")

With workOrderList
    Set Rng = .Find(What:=workOrder.Value, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
    If Not Rng Is Nothing Then
        MsgBox ("Work order number alread exists.")
        ValidateWorkOrder = False
    Else
        ValidateWorkOrder = True
    End If
End With

End Function

Private Function ValidateBoxNumbers(boxStart As String, boxEnd As String) As Boolean
    Dim startNum As Long
    Dim endNum As Long
    startNum = CLng(boxStart)
    endNum = CLng(boxEnd)
    
    If Not (VarType(startNum) = vbLong And VarType(endNum) = vbLong) Then
        ValidateBoxNumbers = False
    End If
    
    If endNum < startNum Then
        ValidateBoxNumbers = False
    Else
        ValidateBoxNumbers = True
    End If
End Function
