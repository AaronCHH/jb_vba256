Attribute VB_Name = "Module1"
Option Explicit

Sub �ѷӵ���()
    Workbooks.Open Filename:="C:\ExcelVBA\��y����.xls"
    Workbooks.Open Filename:="C:\ExcelVBA\�^�y����.xls"
    Windows.Arrange xlArrangeStyleCascade
    MsgBox "��1��: " & Windows(1).Caption & Chr(10) & _
           "��2��: " & Windows(2).Caption
End Sub


