Attribute VB_Name = "Module1"
Option Explicit

Sub ��������ܩM����()
    Dim myWindows As String
    myWindows = ActiveWindow.Caption
    MsgBox "���ç@�ε���!!"
    ActiveWindow.Visible = False
    MsgBox "�A��ܧ@�ε���!!"
    Windows(myWindows).Visible = True
End Sub
