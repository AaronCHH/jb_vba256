Attribute VB_Name = "Module1"
Option Explicit

Sub ��ܤu�@��()
    Dim i As Integer, myName As Integer
    Dim mySheet As Worksheet
    Worksheets(2).Activate
    MsgBox "��ܪ��u�@��: " & ActiveSheet.Name
    Worksheets(1).Select
    Worksheets(3).Select Replace:=False
End Sub

