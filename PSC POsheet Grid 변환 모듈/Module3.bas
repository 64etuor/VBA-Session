Attribute VB_Name = "Module3"
Sub btn_save_Click()

    Call ShowAll

    Sheets("pr_input").Range("A7:U7").Copy

    Dim lastRow As Long
    lastRow = Sheets("att_raw").Cells(Rows.Count, 1).End(xlUp).Row + 1

    i = MsgBox("저장하시겠습니까?", vbYesNo)
    If i = 6 Then


    If Sheets("pr_input").Range("X18").Value <> 0 Then

        Beep
        MsgBox "Error : 잔여공수가 0이 아닙니다."

        Application.CutCopyMode = False
        Exit Sub
    End If

    If Sheets("pr_input").Range("A7").Value = "" Then
        Beep
        MsgBox "Error : 날짜가 입력되지 않았습니다."

        Application.CutCopyMode = False
        Exit Sub
    End If



    Dim rng As Range
    Set rng = Sheets("att_raw").Range("A:A")
    Set rng2 = Sheets("att_raw").Range("B:B")
    Dim dateValue As Variant
    dateValue = Sheets("pr_input").Range("A7").Value
    linevalue = Sheets("pr_input").Range("A11").Value

    If Application.CountIfs(rng, dateValue, rng2, linevalue) > 0 Then

        Beep
        MsgBox "Error : " & dateValue & "의 데이터가 이미 존재합니다."
        Application.CutCopyMode = False
        Exit Sub
    End If

    Sheets("att_raw").AutoFilterMode = False

    Sheets("att_raw").Range("A" & lastRow).PasteSpecial xlPasteValues

    Application.CutCopyMode = False



    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim rngInput As Range, rngOutput As Range
    Dim prlastRow As Long, rowCount As Long
    Dim uniqueID As Long
    Dim blastrow As Long
        
    
    Sheets("prod_raw").AutoFilterMode = False
    
    Set wsInput = Worksheets("pr_input")
    Set wsOutput = Worksheets("prod_raw")
    
    Set rngInput = wsInput.Range("A11:U39")
    
    rowCount = Application.CountA(rngInput.Columns(1))
    
    prlastRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    
    Set rngOutput = wsOutput.Range("B" & prlastRow + 1).Resize(rowCount, 20)
    
    rngOutput.Value = rngInput.Value
    
''
   
    blastrow = wsOutput.Cells(wsOutput.Rows.Count, "B").End(xlUp).Row
  
    Set rng = wsOutput.Range("B1:B" & blastrow)

    lastID = 0

    For Each cell In rng

        If cell.Value <> "" Then

            If cell.Offset(0, -1).Value = "" Then
                
                cell.Offset(0, -1).Value = lastID + 1
                
                lastID = cell.Offset(0, -1).Value
            Else
                
                lastID = cell.Offset(0, -1).Value
            End If
        End If
    Next cell

    Application.CutCopyMode = False
    
    MsgBox "Data has been successfully saved! " & Chr(13) & Chr(13) & dateValue & "의 데이터가 저장되었습니다."
    
        
        End If
        Application.CutCopyMode = False

End Sub

