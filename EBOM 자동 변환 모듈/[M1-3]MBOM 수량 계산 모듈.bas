Attribute VB_Name = "Module4"
Sub MBOMQtyCal()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim parentLevel As Long
    Dim currentLevel As Long
    Dim i As Long
    Dim parentRow As Long
    Dim parentQty As Double
    Dim currentQty As Double
    Dim secondLevel As Long
    Dim thirdLevel As Long
    Dim forthLevel As Long
    Dim fifthLevel As Long
    Dim sixthLevel As Long
    Dim seventhLevel As Long
    Dim eighthLevel As Long
    Dim ninethLevel As Long
    Dim tenthLevel As Long
 
    Dim secondQty As Double
    Dim thirdQty As Double
    Dim forthQty As Double
    Dim fifthQty As Double
    Dim sixthQty As Double
    Dim seventhQty As Double
    Dim eighthQty As Double
    Dim ninethQty As Double
    Dim tenthQty As Double
    
    Set ws = ActiveWorkbook.ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
   
    parentLevel = 1
    secondLevel = 2
    thirdLevel = 3
    forthLevel = 4
    fifthLevel = 5
    sixthLevel = 6
    seventhLevel = 7
    eighthLevel = 8
    ninethLevel = 9
    tenthLevel = 10

    parentRow = 0
    parentQty = 0
    secondQty = 0
    thirdQty = 0
    forthQty = 0
    fifthQty = 0
    sixthQty = 0
    seventhQty = 0
    eighthQty = 0
    ninethQty = 0
    tenthQty = 0
    parentLevel = 1
    

'Parent Qty and Child Qty calculation
    For i = 2 To lastRow
    
    currentLevel = ws.Cells(i, "L").Value
    currentQty = ws.Cells(i, "O").Value
                
        If currentLevel = parentLevel Then
            
            parentLevel = currentLevel
            parentQty = currentQty
            ws.Cells(i, "P").Value = currentQty
            
        ElseIf currentLevel = parentLevel + 1 Then
        
            ws.Cells(i, "P").Value = parentQty * currentQty
            secondLevel = currentLevel
            secondQty = ws.Cells(i, "P").Value

        ElseIf currentLevel = parentLevel + 2 Then
        
            ws.Cells(i, "P").Value = secondQty * currentQty
            thirdLevel = currentLevel
            thirdQty = ws.Cells(i, "P").Value
                       
         ElseIf currentLevel = parentLevel + 3 Then
        
            ws.Cells(i, "P").Value = thirdQty * currentQty
            forthLevel = currentLevel
            forthQty = ws.Cells(i, "P").Value

         ElseIf currentLevel = parentLevel + 4 Then
        
            ws.Cells(i, "P").Value = forthQty * currentQty
            fifthLevel = currentLevel
            fifthQty = ws.Cells(i, "P").Value

        ElseIf currentLevel = parentLevel + 5 Then
        
            ws.Cells(i, "P").Value = fifthQty * currentQty
            sixthLevel = currentLevel
            sixthQty = ws.Cells(i, "P").Value

        ElseIf currentLevel = parentLevel + 6 Then
        
            ws.Cells(i, "P").Value = sixthQty * currentQty
            seventhLevel = currentLevel
            seventhQty = ws.Cells(i, "P").Value
                       
         ElseIf currentLevel = parentLevel + 7 Then
        
            ws.Cells(i, "P").Value = seventhQty * currentQty
            eighthLevel = currentLevel
            eighthQty = ws.Cells(i, "P").Value

         ElseIf currentLevel = parentLevel + 8 Then
        
            ws.Cells(i, "P").Value = eighthQty * currentQty
            ninethLevel = currentLevel
            ninethQty = ws.Cells(i, "P").Value
         
         ElseIf currentLevel = parentLevel + 9 Then
        
            ws.Cells(i, "P").Value = ninethQty * currentQty
            tenthLevel = currentLevel
            tenthQty = ws.Cells(i, "P").Value
        End If
        
    Next i
End Sub

