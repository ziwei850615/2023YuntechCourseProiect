Attribute VB_Name = "Module2"
Sub anomalyHighLight()

Dim targetCol As Long
targetCol = CLng(InputBox("�п�J�A�n���R�������,�d��:�p�G�O��2��N�O�п�J2"))
Dim targetValue As Long
targetValue = CLng(InputBox("�п�J�A���`�е���"))

Dim dtRange As Range

Set dtRange = ActiveSheet.UsedRange

rowNum = dtRange.Rows.Count

Dim rowIdx As Long
'�q�ĤG�C��̫�@�C�i��Ȼs�ƼЦ�B�z
For rowIdx = 2 To rowNum

  '�P�_����C�@�C���ȬO�_�W�L���`��
  If Cells(rowIdx, targetCol).Value > targetValue Then
  
   '�p�G�O
   Cells(rowIdx, targetCol).Font.ColorIndex = 3
   Cells(rowIdx, targetCol).Interior.ColorIndex = 6
  
  Else
  
    '�_�h
     Cells(rowIdx, targetCol).Font.ColorIndex = 1
     Cells(rowIdx, targetCol).Interior.ColorIndex = 0
  
  End If
  


Next
End Sub


