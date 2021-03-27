Attribute VB_Name = "Module1"
Option Explicit

Sub 選擇工作表()
    Dim i As Integer, myName As Integer
    Dim mySheet As Worksheet
    Worksheets(2).Activate
    MsgBox "選擇的工作表: " & ActiveSheet.Name
    Worksheets(1).Select
    Worksheets(3).Select Replace:=False
End Sub

