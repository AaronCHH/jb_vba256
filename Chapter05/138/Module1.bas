Attribute VB_Name = "Module1"
Option Explicit

Sub 指定檔名存檔()
    Dim fPath As String
    fPath = "C:\ExcelVBA\"
    Workbooks.Open Filename:=fPath & "測試結果表單.xls"
    With Application.FileDialog(msoFileDialogSaveAs)
        .FilterIndex = 1
        .InitialFileName = fPath & "互換\結果"
        If .Show = -1 Then .Execute
    End With
End Sub

