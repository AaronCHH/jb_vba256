Attribute VB_Name = "Module1"
Option Explicit

Sub �s�W����()
    On Error GoTo errHandler
    Range("A1").AddComment "������: " & Application.UserName
    With Range("A10")
        .AddComment "��J��: " & Chr(10) & Date
        .Comment.Shape.AutoShapeType = msoShape24pointStar
        .Comment.Visible = True
    End With
    Exit Sub
errHandler:
    MsgBox "�������J����!!"
End Sub

