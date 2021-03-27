Attribute VB_Name = "Module1"
Option Explicit

Sub ?X??A?X}?C()
    Dim myName() As String
    Dim cnt As Integer, i As Integer
    cnt = Range("A1").End(xlDown).Row
    ReDim myName(cnt - 1)
    For i = 0 To cnt - 1
        myName(i) = Cells(i + 1, 1).Value
        Worksheets(i + 2).Name = myName(i)
    Next i
End Sub

