Attribute VB_Name = "Module1"
Option Explicit

Sub 更新存檔並關閉()
    ActiveWorkbook.Close SaveChanges:=True
End Sub

Sub 存檔為無巨集的活頁簿()
    Worksheets("英語測試").Copy
    ActiveWorkbook.Close SaveChanges:=True, _
                         Filename:="C:\ExcelVBA\成績.xlsx"
End Sub

Sub 關閉所有活頁簿()
    Workbooks.Close
End Sub







