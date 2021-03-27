Attribute VB_Name = "Module1"
Option Explicit

Sub 複製活頁簿()
    Dim bkPath As String, bkName As String
    bkPath = "C:\ExcelVBA\BK\"
    bkName = Format(Date, "mm_dd") & "BK.xls"
    Workbooks.Open Filename:="C:\ExcelVBA\測試結果表單.xls"
    ActiveWorkbook.SaveCopyAs Filename:=bkPath & bkName
End Sub








