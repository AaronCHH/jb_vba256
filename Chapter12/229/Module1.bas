Attribute VB_Name = "Module1"
Option Explicit

Sub �R���Ϫ�H�~���ϧ�()
    Dim myShape As Shape
    For Each myShape In ActiveSheet.Shapes
        If myShape.HasChart = msoFalse Then
           myShape.Delete
        End If
    Next
End Sub

