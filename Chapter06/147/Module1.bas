Attribute VB_Name = "Module1"
Option Explicit

Sub 更改檔名並移動()
    On Error GoTo errHandler
    ChDir "C:\ExcelVBA\"
    Name "人才教育.xls" As "人才開發.xls"
    Name "經理.xls" As CurDir & "\BK\經理.xls"
    Exit Sub
errHandler:
    MsgBox "錯誤編號: " & Err.Number & Chr(10) & "錯誤內容: " & Err.Description
End Sub



