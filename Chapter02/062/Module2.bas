Attribute VB_Name = "Module2"
Option Explicit

Sub �R���W��()
    On Error GoTo errHandler
    ActiveWorkbook.Names("�~���ؼ�").Delete
    Exit Sub
errHandler:
    MsgBox "�S���W��!!"
End Sub

Sub �R�������W��()
    Dim xName As Name
    For Each xName In ActiveWorkbook.Names
        xName.Delete
    Next
End Sub

