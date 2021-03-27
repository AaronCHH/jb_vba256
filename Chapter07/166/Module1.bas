Attribute VB_Name = "Module1"
Option Explicit

Sub 隱藏框線()
    Workbooks.Open Filename:="C:\ExcelVBA\收據書.xls"
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
End Sub

Sub 隱藏公式列()
    Application.DisplayFormulaBar = False
End Sub
