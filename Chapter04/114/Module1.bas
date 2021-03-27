Attribute VB_Name = "Module1"
Option Explicit

Sub 工作表數()
    MsgBox "工作表數: " & Worksheets.Count
    Worksheets(Array(1, 3)).Select
    MsgBox "選擇的工作表數: " & ActiveWindow.SelectedSheets.Count
End Sub


