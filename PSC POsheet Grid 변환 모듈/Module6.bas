Attribute VB_Name = "Module6"
Private Sub UserForm_Initialize()


End Sub

Private Sub lb_def_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim i As Long
    i = lb_def.ListIndex
    If i >= 0 Then
        Sheets("pr_input").Range("A1:K1").Value = ListBox1.List(i)
    End If
End Sub

