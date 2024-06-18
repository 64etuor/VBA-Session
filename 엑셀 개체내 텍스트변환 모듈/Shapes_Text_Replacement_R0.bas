Attribute VB_Name = "Module1"
Sub TextReplacement()

    Dim Ws As Worksheet
    Dim shp As Shape
    Dim xFindStr As String
    Dim xReplace As String
    Dim xValue As String
    Dim count As Long
    Dim totalCount As Long

    xFindStr = Application.InputBox("Find:", xTitleId, "", Type:=2)
    xReplace = Application.InputBox("Replace with:", xTitleId, "", Type:=2)

    On Error Resume Next
    For Each Ws In Application.ActiveWorkbook.Worksheets
        For Each shp In Ws.Shapes
            xValue = shp.TextFrame.Characters.Text
            If InStr(xValue, xFindStr) > 0 Then
                count = UBound(Split(xValue, xFindStr)) - LBound(Split(xValue, xFindStr))
                totalCount = totalCount + count
            End If
            shp.TextFrame.Characters.Text = VBA.Replace(xValue, xFindStr, xReplace, 1)
        Next shp
    Next Ws

    MsgBox totalCount & " 개의 단어가 치환되었습니다.", vbInformation

End Sub
