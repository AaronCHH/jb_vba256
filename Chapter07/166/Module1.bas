Attribute VB_Name = "Module1"
Option Explicit

Sub ���îؽu()
    Workbooks.Open Filename:="C:\ExcelVBA\���ڮ�.xls"
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
End Sub

Sub ���ä����C()
    Application.DisplayFormulaBar = False
End Sub
