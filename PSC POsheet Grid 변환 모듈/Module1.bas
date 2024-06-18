Attribute VB_Name = "Module1"
Private Sub UserForm_Initialize()


End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim i As Long
    i = ListBox1.ListIndex
    If i >= 0 Then
        Sheets("pr_input").Range("A1:K1").Value = ListBox1.List(i)
    End If
End Sub

