Attribute VB_Name = "Module1"
Sub userDefineDetect()
Dim rIdx As Long
Dim cIdx As Long
Dim customizedChar As String
Dim customizedCnt As Long '計數器
Dim rowNum As Long
rowNum = Cells(Rows.Count, 1).End(xlUp).Row
Dim colNum As Long
colNum = Cells(1, Columns.Count).End(xlToLeft).Column

customizedCnt = 0
customizedChar = InputBox("請輸入自定義的缺職取代字元,例如-")
For cIdx = 1 To colNum

For rIdx = 1 To rowNum
If Trim(Cells(rIdx, cIdx).Value) Like Trim(customizedChar) Then
customizedCnt = customizedCnt + 1
End If
Next
Next
If customizedCnt > 0 Then
MsgBox "有缺值" & customizedCnt & "筆,需做預處理"

Else
MsgBox "此資料集未缺值儲存格"
End If

End Sub

Sub usedRangedemo()
Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim rowNum As Long
rowNum = dtRange.Rows.Count

Dim colNum As Long
colNum = dtRange.Columns.Count
Dim customizedChar As String
Dim customizedCnt As Long
ustomizedCnt = 0
customizedChar = InputBox("請輸入自定義的缺職取代字元,例如-")
For cIdx = 1 To colNum

For rIdx = 1 To rowNum
If Trim(Cells(rIdx, cIdx).Value) Like Trim(customizedChar) Then
customizedCnt = customizedCnt + 1
End If
Next
Next
If customizedCnt > 0 Then
MsgBox "有缺值" & customizedCnt & "筆,需做預處理"

Else
MsgBox "此資料集未缺值儲存格"
End If

End Sub
