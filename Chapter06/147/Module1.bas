Attribute VB_Name = "Module1"
Option Explicit

Sub ����ɦW�ò���()
    On Error GoTo errHandler
    ChDir "C:\ExcelVBA\"
    Name "�H�~�Ш|.xls" As "�H�~�}�o.xls"
    Name "�g�z.xls" As CurDir & "\BK\�g�z.xls"
    Exit Sub
errHandler:
    MsgBox "���~�s��: " & Err.Number & Chr(10) & "���~���e: " & Err.Description
End Sub



