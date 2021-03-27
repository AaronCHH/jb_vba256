Attribute VB_Name = "Module2"
Option Explicit

Sub 刪除名稱()
    On Error GoTo errHandler
    ActiveWorkbook.Names("年間目標").Delete
    Exit Sub
errHandler:
    MsgBox "沒有名稱!!"
End Sub

Sub 刪除全部名稱()
    Dim xName As Name
    For Each xName In ActiveWorkbook.Names
        xName.Delete
    Next
End Sub

