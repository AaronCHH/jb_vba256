Attribute VB_Name = "Module1"
Option Explicit

Sub �ѷӧ@�Τ�������ï()
    Dim xBook As Variant, i As Integer
    MsgBox "�ثe�@�Τ�������ï: " & ActiveWorkbook.Name & Chr(10) & _
           "����{����������ï: " & ThisWorkbook.Name
    Workbooks.Open Filename:="���D1.xls"
    MsgBox "�ثe�@�Τ�������ï: " & ActiveWorkbook.Name & Chr(10) & _
           "����{����������ï: " & ThisWorkbook.Name
End Sub





