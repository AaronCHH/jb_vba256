Attribute VB_Name = "Module1"
Option Explicit

Sub 設定檔案屬性()
    Dim myFile As String
    myFile = "D:\mee\旗標事務\專案\F0034Excel VBA目的快查式參考手冊\範例檔案\Chapter06\ExcelVBA\測試結果表單.xls"
    SetAttr myFile, vbReadOnly
End Sub

Sub 取得屬性()
    On Error Resume Next
    MsgBox "製作人: " & ActiveWorkbook.BuiltinDocumentProperties("Author") & Chr(10) & _
           "製作時間: " & ActiveWorkbook.BuiltinDocumentProperties("Creation date")
End Sub

Sub 刪除屬性()
    ActiveWorkbook.RemoveDocumentInformation (xlRDIDocumentProperties)
End Sub

