VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frontend 
   Caption         =   "Data Categorizer v. 1.0-ALPHA"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8925
   OleObjectBlob   =   "Frontend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frontend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private active As Range

Private nextCell As Range


Private Sub btnCategorize_Click()
    active.Clear
    active.value = cb_categories.Text
    Set nextCell = active.Offset(1, 0)
    If nextCell.Text = "" Then
        MsgBox "Finished Data Analysis for this row. Exiting."
        Unload Me
        Exit Sub
    End If
    Set active = nextCell
    active.Select
    LoadUserData
    LoadControls
End Sub

Private Sub UserForm_Initialize()
     Backend.InitializeCoroutines
     LoadUserData
     LoadControls
     RefreshDb
End Sub

Private Sub btn_add_Click()
    Backend.AddKVP Backend.FormatKey(cat_name.Text), cat_desc.Text
    ClearControls
    LoadControls
    RefreshDb
End Sub

' Private Sub btn_del_Click()
  '  If Not IsEmpty(cat_name.Text) Then
   '     Backend.RemoveKVP (cat_name.Text)
    '    ClearControls
     '   LoadControls
    'End If
'End Sub

Private Sub cb_categories_Change()
    tb_description.Text = Backend.EnumsData(cb_categories.Text)
End Sub

Private Sub ListBox1_Click()
    cat_name.Text = ListBox1.Text
    cat_desc.Text = Backend.EnumsData(ListBox1.Text)
End Sub


Private Sub LoadControls()
    If Backend.EnumsWorksheet Is Nothing Then
        Exit Sub
    End If
    
    If Backend.EnumsData Is Nothing Then
        Exit Sub
    End If
    
    If Not IsEmpty(active.Text) Then
        TextBox1.Text = active.Text
        lblStatus = "Current Cell : " & active.AddressLocal
    End If
    
End Sub

Private Sub RefreshDb()
    For Each Item In Backend.EnumsData.Keys
        cb_categories.AddItem (Item)
        ListBox1.AddItem (Item)
    Next Item
End Sub

Private Sub ClearControls()
    cb_categories.Clear
    ListBox1.Clear
End Sub

Private Sub LoadUserData()
    Set active = Application.activeCell
    LoadControls
End Sub
