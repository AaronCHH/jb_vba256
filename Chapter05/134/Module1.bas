Attribute VB_Name = "Module1"
Option Explicit

Sub �@�ά���ï��s�s��()
    Workbooks.Open Filename:="C:\ExcelVBA\�ƾ�.xls"
    Range("G1").Value = Date
   'Workbooks(2).Worksheets(1).Range("G1").Value = Date
    ActiveWorkbook.Save
End Sub

Sub ����ï��s�s��()
    Dim xBook As Workbook
    For Each xBook In Workbooks
        xBook.Save
    Next
End Sub
