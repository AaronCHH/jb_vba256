Attribute VB_Name = "Module1"
Option Explicit

Sub Loop1()
    Dim i As Integer
    i = 4
    Do While Cells(i, 1).Value <> ""
       If Cells(i, 3).Value >= 1 Then
          Cells(i, 3).Interior.ColorIndex = 38
       End If
       i = i + 1
    Loop
End Sub


