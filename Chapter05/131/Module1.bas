Attribute VB_Name = "Module1"
Option Explicit

Sub ���}����ï()
    On Error GoTo errHandler
    Workbooks.Open Filename:="�~�Z��.xls", ReadOnly:=True
    Exit Sub
errHandler:
    MsgBox "�䤣����w���ɮ�!!!"
End Sub


