Attribute VB_Name = "Module2"
Option Explicit

Sub CellJump()
    Application.Goto Range("A8:F12")
End Sub

Sub CellJump2()
    Application.Goto Reference:=Worksheets("5ды").Range("A8"), Scroll:=True
End Sub

