Attribute VB_Name = "Module1"
Option Explicit

Sub 複製檔案()
    On Error GoTo errHandler
    FileCopy Source:="C:\ExcelVBA\Data.xls", Destination:="C:\ExcelVBA\Data_BK.xls"
    Exit Sub
errHandler:
    MsgBox "錯誤編號: " & Err.Number & Chr(10) & "錯誤內容: " & Err.Description
End Sub

Sub 刪除檔案()
    On Error GoTo errHandler
    Kill "C:\ExcelVBA\Data_BK.xls"
    Exit Sub
errHandler:
    MsgBox "錯誤編號: " & Err.Number & Chr(10) & "錯誤內容: " & Err.Description
End Sub




