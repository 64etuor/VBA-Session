Attribute VB_Name = "Module7"
Option Explicit

Public netdata As Integer

Sub definput() '공정생산Input
Dim iRows As Long
Dim c, r As Long
Dim vr()
    
    With Sheets("prod_raw").Range("A1")
       
    netdata = .CurrentRegion.Columns.Count
    iRows = .CurrentRegion.Rows.Count

     ReDim vr(iRows - 1, netdata)
     For c = 0 To netdata - 1
          For r = 0 To iRows - 1
            vr(r, c) = .Offset(r, c).Value
         Next r
     Next c
    End With

    With form_def.lb_def
    .Visible = False
    .Clear
    .List() = vr
    .Visible = True
    End With

End Sub


Sub show_def() 'def 결과값 출력
Dim vr()
Dim v
Dim c, r, i As Long
Dim dbrng As Range

 netdata = Sheets("prod_raw").Range("A1").CurrentRegion.Columns.Count
 Set dbrng = Sheets("prod_raw").Range("A1").CurrentRegion.Columns(1).SpecialCells(xlCellTypeVisible)
 
 With dbrng
    For Each v In dbrng
        r = r + 1
        ReDim Preserve vr(1 To netdata, 1 To r)
         
        For i = 1 To netdata
            vr(i, r) = v.Cells(1, i)
        Next i
        c = c + 1
    Next
End With

With form_def.lb_def
    .Visible = False
    .Clear
    .ColumnCount = netdata
    If c = 1 Then
    
    Else
        .List() = Application.Transpose(vr)
    End If
    .Visible = True
End With

End Sub

