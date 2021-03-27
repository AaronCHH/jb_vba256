Attribute VB_Name = "Module1"
Option Explicit

Sub Loop4()
    Dim myRange As Range
    
    For Each myRange In Range("B2", "D7")
        If myRange.Value = "" Then
           myRange.Value = "?????X"
           myRange.Interior.ColorIndex = 35
        End If
    Next
End Sub

