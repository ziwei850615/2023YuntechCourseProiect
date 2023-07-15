Attribute VB_Name = "Module1"
Sub 口罩特約藥局排序練習0714()
Attribute 口罩特約藥局排序練習0714.VB_Description = "目前醫療口罩特約藥局庫存由小到大排序"
Attribute 口罩特約藥局排序練習0714.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 口罩特約藥局排序練習0714 巨集
' 目前醫療口罩特約藥局庫存排序
'
' 快速鍵: Ctrl+a
'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "口罩數量"
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 口罩特約藥局0714練習二()
Attribute 口罩特約藥局0714練習二.VB_Description = "由大到小排序"
Attribute 口罩特約藥局0714練習二.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' 口罩特約藥局0714練習二 巨集
'
' 快速鍵: Ctrl+c
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
