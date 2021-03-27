Attribute VB_Name = "Module1"
Option Explicit

Sub 作用活頁簿更新存檔()
    Workbooks.Open Filename:="C:\ExcelVBA\數學.xls"
    Range("G1").Value = Date
   'Workbooks(2).Worksheets(1).Range("G1").Value = Date
    ActiveWorkbook.Save
End Sub

Sub 活頁簿更新存檔()
    Dim xBook As Workbook
    For Each xBook In Workbooks
        xBook.Save
    Next
End Sub
