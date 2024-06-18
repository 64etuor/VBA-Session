Attribute VB_Name = "Module3"
Sub MBOMFill()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim parentLevel As Long
    Dim currentLevel As Long
    Dim i As Long
    Dim parentRow As Long
    
    
    Set ws = ActiveWorkbook.ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    
    
    parentLevel = 0
    parentRow = 0
    
    
    For i = 2 To lastRow
            If ws.Cells(i, "R").Value = "Buy" Then
            parentRow = i
            parentLevel = ws.Cells(i, "L").Value
            Exit For
        End If
    Next i
    
    
    If parentRow = 0 Then
        MsgBox """BOM.Buy"" row not found in the specified column."
        Exit Sub
    End If
    
    
    For i = parentRow To lastRow
    
        currentLevel = ws.Cells(i, "L").Value
        
        If ws.Cells(i, "R").Value = "Buy" And currentLevel <= parentLevel Then
            
            ws.Cells(i, "S").Value = "E"
                        
            parentLevel = currentLevel
        
        ElseIf ws.Cells(i, "R").Value = "Make" And currentLevel <= parentLevel Then
            
            parentLevel = currentLevel + 1
            
        ElseIf currentLevel > parentLevel And ws.Cells(i, "R").Value = "Buy" Then
     
            ws.Cells(i, "S").Value = "EP"
            
        End If
    Next i
    
Application.CutCopyMode = False
End Sub

