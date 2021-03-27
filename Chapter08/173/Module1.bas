Attribute VB_Name = "Module1"
Option Explicit

Sub 預覽列印()
    Dim x As Integer
    x = Application.InputBox(Prompt:="請選擇預覽列印範圍" & _
        Chr(10) & _
        "1: 3教科工作表預覽" & Chr(10) & _
        "2: 各表格單位預覽" & Chr(10) & _
        "3: 科目別工作表預覽", Type:=1)
    Select Case x
        Case 1: ActiveSheet.PrintPreview False
        Case 2: ActiveSheet.Range("A1:F14,A17:F30,A33:G46").PrintPreview
        Case 3: Worksheets(Array(2, 3, 4)).PrintPreview False
        Case Else: Exit Sub
    End Select
End Sub

