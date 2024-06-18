Attribute VB_Name = "Module8"
Sub 불량시트필터초기화()
Attribute 불량시트필터초기화.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 불량시트필터초기화 매크로
'

'
    ActiveWorkbook.SlicerCaches("슬라이서_불량유형_1").ClearManualFilter
    ActiveWorkbook.SlicerCaches("슬라이서_PartName").ClearManualFilter
    ActiveWorkbook.SlicerCaches("슬라이서_Process").ClearManualFilter
    ActiveWorkbook.SlicerCaches("슬라이서_Weeknum").ClearManualFilter
    ActiveWorkbook.SlicerCaches("슬라이서_Line").ClearManualFilter

End Sub
