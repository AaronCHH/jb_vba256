Attribute VB_Name = "Module1"
Option Explicit

Sub Loop2()
    Dim i As Integer
    i = 4
    Do Until Month(Cells(i, 1).Value) = 4
       If Cells(i, 3).Value >= 1 Then
          Cells(i, 3).Interior.ColorIndex = 38
       End If
       i = i + 1
    Loop
End Sub

