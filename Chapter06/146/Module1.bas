Attribute VB_Name = "Module1"
Option Explicit

Sub �ƻs�ɮ�()
    On Error GoTo errHandler
    FileCopy Source:="C:\ExcelVBA\Data.xls", Destination:="C:\ExcelVBA\Data_BK.xls"
    Exit Sub
errHandler:
    MsgBox "���~�s��: " & Err.Number & Chr(10) & "���~���e: " & Err.Description
End Sub

Sub �R���ɮ�()
    On Error GoTo errHandler
    Kill "C:\ExcelVBA\Data_BK.xls"
    Exit Sub
errHandler:
    MsgBox "���~�s��: " & Err.Number & Chr(10) & "���~���e: " & Err.Description
End Sub




