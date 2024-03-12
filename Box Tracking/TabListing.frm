VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TabListing 
   Caption         =   "Tab Listing"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5265
   OleObjectBlob   =   "TabListing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TabListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hideButton_Click()
    For i = 0 To TabListing.tabListBox.ListCount - 1
        If TabListing.tabListBox.selected(i) Then
            ActiveWorkbook.Worksheets(TabListing.tabListBox.list(i)).Visible = False
        End If
    Next i
    
    Call PopulateTabListing
End Sub

Private Sub showButton_Click()
    For i = 0 To TabListing.tabListBox.ListCount - 1
        If TabListing.tabListBox.selected(i) Then
            ActiveWorkbook.Worksheets(TabListing.tabListBox.list(i)).Visible = True
        End If
    Next i

    Call PopulateTabListing
End Sub
