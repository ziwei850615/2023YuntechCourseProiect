Attribute VB_Name = "Module1"
Sub �f�n�S���ħ��Ƨǽm��0714()
Attribute �f�n�S���ħ��Ƨǽm��0714.VB_Description = "�ثe�����f�n�S���ħ��w�s�Ѥp��j�Ƨ�"
Attribute �f�n�S���ħ��Ƨǽm��0714.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' �f�n�S���ħ��Ƨǽm��0714 ����
' �ثe�����f�n�S���ħ��w�s�Ƨ�
'
' �ֳt��: Ctrl+a
'
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "�f�n�ƶq"
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub �f�n�S���ħ�0714�m�ߤG()
Attribute �f�n�S���ħ�0714�m�ߤG.VB_Description = "�Ѥj��p�Ƨ�"
Attribute �f�n�S���ħ�0714�m�ߤG.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' �f�n�S���ħ�0714�m�ߤG ����
'
' �ֳt��: Ctrl+c
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
