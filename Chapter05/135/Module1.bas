Attribute VB_Name = "Module1"
Option Explicit

Sub ����ï�t�s�s��()
    Dim fPath As String, fName As String
    fPath = "C:\ExcelVBA\"
    Workbooks.Open Filename:=fPath & "���յ��G���.xls"
    fName = Format(Date, "mm_dd") & "���G"
    
    ActiveWorkbook.Worksheets("���G").Name = fName
    ActiveWorkbook.SaveAs _
        Filename:=fPath & fName
End Sub







