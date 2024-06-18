Attribute VB_Name = "Module5"
Option Explicit

Public netdata As Integer
Public Fileobj As Object
Public Path As String
Public FileName As String
Public wost As Worksheet


Sub woinput() '워크오더Input
Dim iRows As Long
Dim c, r As Long
Dim vr()
    
Path = ThisWorkbook.Path & "\db\"
FileName = "prod_raw.xlsx"
Set Fileobj = GetObject(Path + FileName)
Set wost = Fileobj.Sheets("wo_raw")
    
    With wost.Range("A1")
       
    netdata = .CurrentRegion.Columns.Count
    iRows = .CurrentRegion.Rows.Count

     ReDim vr(iRows - 1, netdata)
     For c = 0 To netdata - 1
          For r = 0 To iRows - 1
            vr(r, c) = .Offset(r, c).Value
         Next r
     Next c
    End With

    With UserForm1.ListBox1
    .Visible = False
    .Clear
    .List() = vr
    .Visible = True
    End With

End Sub

Sub fdinput() '필터 데이터 Input

Dim c As Long
Dim vr()

Path = ThisWorkbook.Path & "\db\"
FileName = "prod_raw.xlsx"
Set Fileobj = GetObject(Path + FileName)
Set wost = Fileobj.Sheets("wo_raw")
    

      With wost.Range("A1")
         netdata = .CurrentRegion.Columns.Count - 1
         ReDim vr(netdata)
            For c = 0 To netdata
            vr(c) = .Cells(1, c + 1).Value
       Next c
    End With
    
    With UserForm1.ComboBox1
        .Visible = False
        .Clear
        .ListRows = netdata + 1
        .List() = vr
        .Visible = True
    End With

    
End Sub

Sub showresult() '결과값 출력
Dim vr()
Dim v
Dim c, r, i As Long
Dim dbrng As Range

Path = ThisWorkbook.Path & "\db\"
FileName = "prod_raw.xlsx"
Set Fileobj = GetObject(Path + FileName)
Set wost = Fileobj.Sheets("wo_raw")
    

 netdata = wost.Range("A1").CurrentRegion.Columns.Count
 Set dbrng = wost.Range("A1").CurrentRegion.Columns(1).SpecialCells(xlCellTypeVisible)
 
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

With UserForm1.ListBox1
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
