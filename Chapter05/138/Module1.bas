Attribute VB_Name = "Module1"
Option Explicit

Sub ���w�ɦW�s��()
    Dim fPath As String
    fPath = "C:\ExcelVBA\"
    Workbooks.Open Filename:=fPath & "���յ��G���.xls"
    With Application.FileDialog(msoFileDialogSaveAs)
        .FilterIndex = 1
        .InitialFileName = fPath & "����\���G"
        If .Show = -1 Then .Execute
    End With
End Sub

