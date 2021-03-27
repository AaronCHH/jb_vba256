Attribute VB_Name = "Module1"
Option Explicit

Sub 打開活頁簿()
    On Error GoTo errHandler
    Workbooks.Open Filename:="業績表.xls", ReadOnly:=True
    Exit Sub
errHandler:
    MsgBox "找不到指定的檔案!!!"
End Sub


