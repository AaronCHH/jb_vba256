Attribute VB_Name = "Module1"
Option Explicit

Sub 新增註解()
    On Error GoTo errHandler
    Range("A1").AddComment "做成者: " & Application.UserName
    With Range("A10")
        .AddComment "輸入日: " & Chr(10) & Date
        .Comment.Shape.AutoShapeType = msoShape24pointStar
        .Comment.Visible = True
    End With
    Exit Sub
errHandler:
    MsgBox "完成插入註解!!"
End Sub

