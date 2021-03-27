Attribute VB_Name = "Module1"
Option Explicit

Sub 活頁簿另存新檔()
    Dim fPath As String, fName As String
    fPath = "C:\ExcelVBA\"
    Workbooks.Open Filename:=fPath & "測試結果表單.xls"
    fName = Format(Date, "mm_dd") & "結果"
    
    ActiveWorkbook.Worksheets("結果").Name = fName
    ActiveWorkbook.SaveAs _
        Filename:=fPath & fName
End Sub







