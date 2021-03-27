Attribute VB_Name = "Module1"
Option Explicit

Sub ¬d¸ß¦r¦ê()
    Dim i As Integer, s As Integer
    Columns("B:C").NumberFormatLocal = "@"
    i = 2
    Do Until Cells(i, 1) = ""
       s = InStr(Cells(i, 1), "-")
       Cells(i, 2).Value = Left(Cells(i, 1), s - 1)
       Cells(i, 3).Value = Mid(Cells(i, 1), s + 1)
       i = i + 1
    Loop
End Sub



