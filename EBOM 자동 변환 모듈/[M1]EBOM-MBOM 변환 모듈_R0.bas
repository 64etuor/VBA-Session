Attribute VB_Name = "Module1"

Sub EBOMRenew()
Attribute EBOMRenew.VB_ProcData.VB_Invoke_Func = "D\n14"

    Dim ws As Worksheet
    Dim i As Long
        
    Set ws = ActiveWorkbook.ActiveSheet
    
 
    For i = 1 To 11
        ws.Columns(1).Insert Shift:=xlToRight
        
    Next i
    
    For i = 1 To 11
    ws.Cells(1, i).Value = i - 1
    
    Next i
    
    ws.Columns("A:K").ColumnWidth = 2.5
       
    Dim find As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
'Level cut and paste
    
    Set find = ws.Rows(1).find("Level", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("L").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("L")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'Level' not found!", vbExclamation, "Error"
    End If
    
'Number cut and paste
    
    Set find = ws.Rows(1).find("Number", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("M").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("M")
        
        Application.CutCopyMode = False
        
    Else
        MsgBox "Column 'Number' not found!", vbExclamation, "Error"
    End If
    
'Product Code 추가

    ws.Columns(14).Insert Shift:=xlToRight
    
    ws.Cells(1, 14).Value = "Product Code"
    
    lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
    
    For r = 2 To lastRow

        ws.Cells(r, 14).Value = ws.Cells(2, 13)

    Next r
    
'BOM.Qty cut and paste
    
    Set find = ws.Rows(1).find("BOM.Qty", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        
        ws.Columns("O").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("O")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'BOM.Qty' not found!", vbExclamation, "Error"
    End If

'M-BOM.Qty 열 추가

    ws.Columns(16).Insert Shift:=xlToRight
    
    ws.Cells(1, 16).Value = "M-BOM.Qty"
    
'BOM.UOM cut and paste
    
    Set find = ws.Rows(1).find("BOM.UOM", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("Q").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("Q")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'BOM.UOM' not found!", vbExclamation, "Error"
    End If
    
'BOM.Buy/Make cut and paste
    
    Set find = ws.Rows(1).find("BOM.Buy/Make", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("R").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("R")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'BOM.Buy/Make' not found!", vbExclamation, "Error"
    End If
        
 'M-BOM.E/EP/SC 열 추가
 
    ws.Columns(19).Insert Shift:=xlToRight
    
    ws.Cells(1, 19).Value = "M-BOM.E/EP/SC"

'Description cut and paste
    
    Set find = ws.Rows(1).find("Description", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("T").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("T")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'Description' not found!", vbExclamation, "Error"
    End If
    
    ws.Columns("T").ColumnWidth = 30
    
 'BOM.Subsidiary Companies Parts cut and paste
    
    Set find = ws.Rows(1).find("BOM.Subsidiary Companies Parts", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("U").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("U")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'BOM.Subsidiary Companies Parts' not found!", vbExclamation, "Error"
    End If
    
 'Manufacturers.Mfr. Part Number cut and paste
    
    Set find = ws.Rows(1).find("Manufacturers.Mfr. Part Number", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("V").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("V")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'Manufacturers.Mfr. Part Number' not found!", vbExclamation, "Error"
    End If
    
 'Part Type cut and paste
    
    Set find = ws.Rows(1).find("Part Type", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("W").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("W")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'Part Type' not found!", vbExclamation, "Error"
    End If
    
 'Manufacturers.Mfr. Name cut and paste
    
    Set find = ws.Rows(1).find("Manufacturers.Mfr. Name", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("X").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("X")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'Manufacturers.Mfr. Name' not found!", vbExclamation, "Error"
    End If
        
  'BOM.Item Description cut and paste
    
    Set find = ws.Rows(1).find("BOM.Item Description", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("Y").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("Y")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'BOM.Item Description' not found!", vbExclamation, "Error"
    End If
    
'z부터 빈 열 6개 추가

    For v = 1 To 6
         
        ws.Columns("Z").Insert Shift:=xlToRight
        
    Next v
    
  'Manufacturers.Preferred Status cut and paste
    
    Set find = ws.Rows(1).find("Manufacturers.Preferred Status", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not find Is Nothing Then
        
        lastRow = ws.Cells(ws.Rows.Count, find.Column).End(xlUp).Row
        
        ws.Columns("AF").Insert Shift:=xlToRight
        
        find.EntireColumn.Cut Destination:=ws.Columns("AF")
        
        Application.CutCopyMode = False
                    
        
    Else
        MsgBox "Column 'Manufacturers.Preferred Status' not found!", vbExclamation, "Error"
    End If

    Set find = Nothing
    

'AG부터 제거

ws.Columns("AG:DT").Clear

'0 ~ 10 열에 숫자 채우기
Dim x As Long
Dim y As Long
     
    
    lastRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    
 
    For x = 2 To lastRow
        For y = 1 To 11
    
    If ws.Cells(x, 12).Value = y - 1 Then
        ws.Cells(x, y).Value = ws.Cells(x, 12).Value
    End If
    
        Next y
        
    Next x

'열 너비 조정

ws.Columns("L").ColumnWidth = 6.43
ws.Columns("M:N").ColumnWidth = 20
ws.Columns("U:V").ColumnWidth = 25
ws.Columns("X:Y").ColumnWidth = 25

'로켈 넘버 변환

Call NumVal

'M BOM 수량 계산 모듈 불러오기

Call MBOMQtyCal

'M BOM E/EP/SC 모듈불러오기

Call MBOMFill

ActiveSheet.UsedRange.AutoFilter

ws.Cells(1, 1).Select

End Sub

