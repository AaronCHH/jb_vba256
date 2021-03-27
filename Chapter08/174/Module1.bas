Attribute VB_Name = "Module1"
Option Explicit

Sub 列印()
    Dim x As Integer
    x = Application.InputBox(Prompt:="請選擇預覽列印範圍" & _
        Chr(10) & _
        "1: [3教科]工作表: 全部" & Chr(10) & _
        "2: [3教科]工作表: 限表格" & Chr(10) & _
        "3: [3教科]工作表: 限圖表" & Chr(10) & _
        "4: 活頁簿的所有工作表", Type:=1)
    Select Case x
        Case 1: ActiveSheet.PrintOut Preview:=True
        Case 2: ActiveSheet.Range("A1:F14").PrintOut Preview:=True
        Case 3: ActiveSheet.ChartObjects(1).Chart.PrintOut Preview:=True
        Case 4: ActiveWorkbook.PrintOut Preview:=True
        Case Else: Exit Sub
    End Select
End Sub



