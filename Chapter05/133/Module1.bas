Attribute VB_Name = "Module1"
Option Explicit

Sub ��s�s�ɨ�����()
    ActiveWorkbook.Close SaveChanges:=True
End Sub

Sub �s�ɬ��L����������ï()
    Worksheets("�^�y����").Copy
    ActiveWorkbook.Close SaveChanges:=True, _
                         Filename:="C:\ExcelVBA\���Z.xlsx"
End Sub

Sub �����Ҧ�����ï()
    Workbooks.Close
End Sub







