Attribute VB_Name = "Module1"
Option Explicit

Sub 配合選擇範圍調整比率()
    Workbooks.Open Filename:="C:\ExcelVBA\收據書.xls"
    Range("A1:E17").Select
    ActiveWindow.Zoom = True
    MsgBox "現在的倍率: " & ActiveWindow.Zoom & "%"
    ActiveWindow.Zoom = 100
End Sub


