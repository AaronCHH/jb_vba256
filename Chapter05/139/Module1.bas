Attribute VB_Name = "Module1"
Option Explicit

Sub �ƻs����ï()
    Dim bkPath As String, bkName As String
    bkPath = "C:\ExcelVBA\BK\"
    bkName = Format(Date, "mm_dd") & "BK.xls"
    Workbooks.Open Filename:="C:\ExcelVBA\���յ��G���.xls"
    ActiveWorkbook.SaveCopyAs Filename:=bkPath & bkName
End Sub








