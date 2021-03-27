Attribute VB_Name = "Module1"
Option Explicit

Sub 刪除圖表以外的圖形()
    Dim myShape As Shape
    For Each myShape In ActiveSheet.Shapes
        If myShape.HasChart = msoFalse Then
           myShape.Delete
        End If
    Next
End Sub

