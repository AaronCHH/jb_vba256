Attribute VB_Name = "Module1"
Option Explicit

Sub 取得簿名測試()
    With Application.FileDialog(msoFileDialogOpen)
        .FilterIndex = 2
        If .Show = 0 Then Exit Sub
        .Execute
    End With
    MsgBox "Name: " & ActiveWorkbook.Name & Chr(10) & _
           "FullName: " & ActiveWorkbook.FullName
End Sub





