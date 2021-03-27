Attribute VB_Name = "Module1"
Option Explicit

Sub 取得部分字串()
    Dim i As Integer
    For i = 2 To 6
        Cells(i, 2) = "'" & Left(Cells(i, 1), 2)
        Cells(i, 3) = "'" & Mid(Cells(i, 1), 3, 4)
        Cells(i, 4) = "'" & Right(Cells(i, 1), 2)
    Next
End Sub



