Attribute VB_Name = "Module1"
Option Explicit

Sub ½Æ»s()
    Range("A5:F9").Copy
    ActiveSheet.Paste Range("A13")
    ActiveSheet.Paste Range("A21")
    Application.CutCopyMode = False
End Sub

