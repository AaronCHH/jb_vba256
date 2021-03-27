Attribute VB_Name = "Module1"
Option Explicit

Sub 活頁簿的更新存檔確認()
    If ActiveWorkbook.Saved Then
        MsgBox "活頁簿不需存檔!!!"
    Else
        MsgBox "已修改!!更新存檔!!"
        ActiveWorkbook.Save
    End If
End Sub

Sub 所有的活頁簿不存檔直接關閉()
    Dim xBook As Workbook
    For Each xBook In Workbooks
        xBook.Saved = True
        xBook.Close
    Next
End Sub


