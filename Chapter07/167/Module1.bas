Attribute VB_Name = "Module1"
Option Explicit

Sub �t�X��ܽd��վ��v()
    Workbooks.Open Filename:="C:\ExcelVBA\���ڮ�.xls"
    Range("A1:E17").Select
    ActiveWindow.Zoom = True
    MsgBox "�{�b�����v: " & ActiveWindow.Zoom & "%"
    ActiveWindow.Zoom = 100
End Sub


