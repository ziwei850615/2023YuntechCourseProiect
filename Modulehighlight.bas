Attribute VB_Name = "Module2"
Sub anomalyHighLight()

Dim targetCol As Long
targetCol = CLng(InputBox("請輸入你要分析的欄索引,範例:如果是第2行就是請輸入2"))
Dim targetValue As Long
targetValue = CLng(InputBox("請輸入你異常標註值"))

Dim dtRange As Range

Set dtRange = ActiveSheet.UsedRange

rowNum = dtRange.Rows.Count

Dim rowIdx As Long
'從第二列到最後一列進行客製化標住處理
For rowIdx = 2 To rowNum

  '判斷該欄每一列的值是否超過異常值
  If Cells(rowIdx, targetCol).Value > targetValue Then
  
   '如果是
   Cells(rowIdx, targetCol).Font.ColorIndex = 3
   Cells(rowIdx, targetCol).Interior.ColorIndex = 6
  
  Else
  
    '否則
     Cells(rowIdx, targetCol).Font.ColorIndex = 1
     Cells(rowIdx, targetCol).Interior.ColorIndex = 0
  
  End If
  


Next
End Sub


