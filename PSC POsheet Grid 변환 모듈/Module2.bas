Attribute VB_Name = "Module2"
Public Fileobj As Object
Public Path As String
Public FileName As String
Public wost As Worksheet

Sub btn_wo_Click()
UserForm1.Show
End Sub

Sub btn_showdef_click()
form_def.Show

End Sub



Sub ShowAll()
Path = ThisWorkbook.Path & "\db\"
FileName = "prod_raw.xlsx"
Set Fileobj = GetObject(Path + FileName)
Set wost = Fileobj.Sheets("wo_raw")

    On Error Resume Next
    With wost.ListObjects(wost)
        If .AutoFilterMode Then
            .AutoFilter.ShowAllData
        End If
    End With
    On Error GoTo 0
End Sub
