Attribute VB_Name = "Backend"
Public EnumsWorksheet As Worksheet
Public EnumsData As Object


Sub InitializeCoroutines()



If WorksheetExists("Enums") = False Then
    Dim aws As Worksheet
    Dim ac As Range
    Set aws = ActiveSheet
    Set ac = Application.activeCell
    ActiveWorkbook.Sheets.Add(After:=ActiveSheet).Name = "Enums"
    aws.Activate
    ac.Select
End If
Set EnumsWorksheet = ActiveWorkbook.Sheets("Enums")
Set EnumsData = CreateObject("Scripting.Dictionary")
ReadKvp
End Sub

Sub RemoveKVP(key As String)
If Not EnumsData.Exists(FormatKey(key)) Then
    EnumsData.Remove FormatKey(key)
    SaveKvp
End If
End Sub

Sub AddKVP(key As String, value As String)
If Not EnumsData.Exists(FormatKey(key)) Then
    EnumsData.Add FormatKey(key), value
    Else
    If EnumsData(FormatKey(key)) <> value Then
        EnumsData(FormatKey(key)) = value
    End If
End If
SaveKvp
End Sub
Sub ReadKvp()
    If EnumsWorksheet Is Nothing Then
        Exit Sub
    End If
    
    If EnumsData Is Nothing Then
        Exit Sub
    End If
    
    If IsEmpty(EnumsWorksheet.Range("A1").value) Then
        Exit Sub
    End If
    
    Dim temp_k, temp_v As Range
    Dim i, nexti As Integer
    Dim kstr, vstr As String
    
    
    
    Set temp_k = EnumsWorksheet.Range("A1").End(xlDown) ' endk [An]
    Set temp_v = EnumsWorksheet.Range("B1").End(xlDown) ' endv [Bn]
    If Not IsEmpty(EnumsWorksheet.Range("A2").value) Then
        Set temp_k = EnumsWorksheet.Range("A1", temp_k.Address) ' datak [A1:An]
        Set temp_v = EnumsWorksheet.Range("B1", temp_v.Address) ' datav [B1:Bn]
    End If
    For i = 0 To temp_k.Count
        Dim temp_ik, temp_iv As Range
        nexti = i + 1
        kstr = "A" & nexti
        vstr = "B" & nexti
        Set temp_ik = EnumsWorksheet.Range(kstr)
        Set temp_iv = EnumsWorksheet.Range(vstr)
        
        If Not IsEmpty(temp_ik.value) Then
            If Not IsEmpty(temp_iv.value) Then
                AddKVP FormatKey(temp_ik.value), temp_iv.value
            End If
        End If
    Next i
    SaveKvp
End Sub

Function FormatKey(str As String) As String
    FormatKey = Replace(UCase(Replace(Replace(Replace(str, ",", " "), ".", " "), "-", " ")), " ", "_")
End Function

Sub SaveKvp()
    If EnumsWorksheet Is Nothing Then
        Exit Sub
    End If
    
    If EnumsData Is Nothing Then
        Exit Sub
    End If
  
    ClearDataset
    Dim i As Integer
    Dim str As String
    Dim nexti As Integer
    i = 0
    For Each key In EnumsData.Keys
        
        nexti = i + 1
        str = "A" & nexti
        EnumsWorksheet.Range(str).value = key
        i = i + 1
    Next key
    i = 0
     For Each key In EnumsData.Items
        nexti = i + 1
        str = "B" & nexti
        EnumsWorksheet.Range(str).value = key
        i = i + 1
    Next key
End Sub

Sub ClearDataset()
    If EnumsWorksheet Is Nothing Then
        Exit Sub
    End If
    
    EnumsWorksheet.Cells.Clear
End Sub


Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
