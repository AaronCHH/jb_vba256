Attribute VB_Name = "Module1"
Option Explicit

Sub �u�@���()
    MsgBox "�u�@���: " & Worksheets.Count
    Worksheets(Array(1, 3)).Select
    MsgBox "��ܪ��u�@���: " & ActiveWindow.SelectedSheets.Count
End Sub


