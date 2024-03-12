Attribute VB_Name = "Clockify_Reporting"
Sub CreateClientReports()
    Dim wb As Workbook
    Dim dataSheet As Worksheet
    
    Dim template As Worksheet
    Dim data As ListObject
    
    Dim clientArray As Variant, uniqueArray As Variant
    Dim i As Integer
    Dim clientDict As Object
    
    Dim maxDate As String
    Dim monthNum As Integer
    Dim month As String
    Dim destinationFolder As String
    
    Set wb = ThisWorkbook
    Set dataSheet = Sheets("DATA")
    Set data = dataSheet.ListObjects("DATA")
    Set template = wb.Sheets("template")
    
    If WorksheetFunction.CountA(Range("Start_Date")) = 0 Then
        Debug.Print ("No data")
        Exit Sub
    End If
    
    maxDate = Format(Application.WorksheetFunction.Max(Range("Start_Date")), "YYYYMMDD")
    monthNum = CInt(Format(Application.WorksheetFunction.Max(Range("Start_Date")), "M"))
    month = MonthName(monthNum)
    destinationFolder = CurDir() & "\" & "Clockify Reporting, " & month
    CreateFolder (destinationFolder)
    
    wb.SaveAs Filename:=destinationFolder & "\_Clockify Full Hours Report, " & month & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    Set clientDict = CreateObject("scripting.dictionary")
    clientArray = Range("Clients").Value
    For i = LBound(clientArray) To UBound(clientArray)
        If (Not clientDict.exists(CStr(clientArray(i, 1)))) Then clientDict.Add CStr(clientArray(i, 1)), clientArray(i, 1)
    Next i
    uniqueArray = clientDict.Items
    
    Application.ScreenUpdating = False
    With data.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("DATA[Client]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range("DATA[Project]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range("DATA[Start Date]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Range("DATA[Task]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    data.ShowAutoFilter = True
    
    'On Error Resume Next
    
    For i = 0 To UBound(uniqueArray)
        Call CreateSingleReport(template, CStr(uniqueArray(i)), data, month, destinationFolder)
    Next i
    Application.ScreenUpdating = True
    
    response = MsgBox(Prompt:="All reports created in the following directory:" & vbNewLine & destinationFolder, _
        Buttons:=vbInformation, _
        Title:="Report Creation Complete!")
    wb.Close SaveChanges:=False
End Sub

Sub CreateSingleReport(template As Worksheet, clientName As String, sourceData As ListObject, month As String, folderPath As String)
    template.Copy Before:=ThisWorkbook.Sheets("DATA")
    With ActiveSheet
        .Name = "HoursReport"
        .Move
    End With
    
    sourceData.Range.AutoFilter Field:=2, Criteria1:=clientName
    'sourceData.DataBodyRange.Copy ActiveSheet.Range("A2")
    sourceData.DataBodyRange.Copy
    ActiveSheet.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    Dim fullpath As String
    With ActiveWorkbook
        fullpath = folderPath & "\" & clientName & " - Clockify Hours, " & month & ".xlsx"
        .SaveAs Filename:=fullpath
        .Close SaveChanges:=False
    End With
    
    Debug.Print ("Created: " & clientName)
End Sub

Sub CreateFolder(path As String)
    If Not (FolderExists(path)) Then MkDir (path)
End Sub

Public Function FolderExists(path As String) As Boolean
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = oFSO.FolderExists(path)
End Function
