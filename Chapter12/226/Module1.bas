Attribute VB_Name = "Module1"
Option Explicit

Sub �ѷӹϧ�()
    Dim i As Integer
    For i = 1 To ActiveSheet.Shapes.Count
        ActiveSheet.Shapes(i).Select
        Selection.Text = i & ": " & Selection.Name
    Next
End Sub


