Attribute VB_Name = "Module2"
Sub NumVal()

    Columns("L").Select
    Selection.NumberFormatLocal = "G/ǥ��"
    Selection.Value = Selection.Value
  
Application.CutCopyMode = False
  
    Columns("O").Select
    Selection.NumberFormatLocal = "G/ǥ��"
    Selection.Value = Selection.Value

Application.CutCopyMode = False

End Sub

