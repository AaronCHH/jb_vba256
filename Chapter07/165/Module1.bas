Attribute VB_Name = "Module1"
Option Explicit

Sub ���ñ��b()
    Workbooks.Open Filename:="C:\ExcelVBA\3�Ь����.xls"
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
End Sub

Sub ���éҦ����������b()
    Application.DisplayScrollBars = Not Application.DisplayScrollBars
End Sub


