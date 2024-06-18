Attribute VB_Name = "Module1"
'Sub OptimizedTransformTable()
'    Dim ws As Worksheet
'    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
'    Dim newWs As Worksheet, newRow As Long
'    Dim dataArray As Variant
'
'    Application.ScreenUpdating = False ' Turn off screen updating to speed up the macro
'    Application.Calculation = xlCalculationManual ' Turn off automatic calculations
'
'    Set ws = ThisWorkbook.Sheets("Original") ' Assuming the data is in a sheet named "Original"
'    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
'    newWs.Name = "Transformed"
'
'    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row
'    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Find the last column
'
'    ' Copy headers and adjust for the transformed table
'    newWs.Range(newWs.Cells(1, 1), newWs.Cells(1, 10)).Value = ws.Range(ws.Cells(1, 1), ws.Cells(1, 10)).Value
'    newWs.Cells(1, 11).Value = "Date" ' Add Date header
'    newWs.Cells(1, 12).Value = "PlanQ'ty" ' Add PlanQ'ty header
'
'    newRow = 2 ' Start from the second row in the new sheet
'    dataArray = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value ' Read the entire range into an array for faster processing
'
'    For i = 2 To UBound(dataArray, 1) ' Loop through each row
'        For j = 12 To UBound(dataArray, 2) ' Loop through each date column
'            If dataArray(i, j) <> "" Then ' Check if there's a quantity
'                ' Copy the first 11 columns
'                newWs.Range(newWs.Cells(newRow, 1), newWs.Cells(newRow, 11)).Value = Application.Index(dataArray, i, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
'                ' Copy the date from the header and the quantity
'                newWs.Cells(newRow, 11).Value = dataArray(1, j)
'                newWs.Cells(newRow, 12).Value = dataArray(i, j)
'                newRow = newRow + 1 ' Move to the next row in the new sheet
'            End If
'        Next j
'    Next i
'
'    newWs.Columns("K:K").Select
'
'    Selection.NumberFormat = "yyyy-mm-dd"
'
'    Application.ScreenUpdating = True ' Turn on screen updating
'    Application.Calculation = xlCalculationAutomatic ' Turn on automatic calculations
'End Sub
'
Sub TransformTable()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim newWs As Worksheet, newRow As Long
    Dim dataArray As Variant
    Dim planId As Long
    
    Application.ScreenUpdating = False ' Turn off screen updating
    Application.Calculation = xlCalculationManual ' Turn off automatic calculations

    Set ws = ThisWorkbook.Sheets("Original") ' Sheet with original data
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "Transformed"
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Find the last row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Find the last column

    ' Copy headers and add new headers
    newWs.Range(newWs.Cells(1, 1), newWs.Cells(1, 10)).Value = ws.Range(ws.Cells(1, 1), ws.Cells(1, 10)).Value
    newWs.Cells(1, 11).Value = "Date"
    newWs.Cells(1, 12).Value = "PlanQ'ty"
    newWs.Cells(1, 13).Value = "Plan_ID" ' Add Plan_ID header

    newRow = 2 ' Starting row for the new sheet
    dataArray = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value ' Array for data
    planId = 1 ' Starting Plan ID

    For i = 2 To UBound(dataArray, 1) ' Loop through rows
        For j = 12 To UBound(dataArray, 2) ' Loop through date columns
            If dataArray(i, j) <> "-" And dataArray(i, j) <> "" Then ' Check for non-empty and non-dash values
                ' Copy the first 11 columns
                newWs.Range(newWs.Cells(newRow, 1), newWs.Cells(newRow, 11)).Value = Application.Index(dataArray, i, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
                newWs.Cells(newRow, 11).Value = dataArray(1, j)
                newWs.Cells(newRow, 12).Value = dataArray(i, j)
                newWs.Cells(newRow, 13).Value = "P" & planId ' Assigning unique Plan_ID
                newRow = newRow + 1
                planId = planId + 1 ' Increment Plan ID
            End If
        Next j
    Next i
    
    newWs.Columns("K:K").Select
    Selection.NumberFormat = "yyyy-mm-dd"

    Application.ScreenUpdating = True ' Turn on screen updating
    Application.Calculation = xlCalculationAutomatic ' Turn on automatic calculations
End Sub

