Attribute VB_Name = "Module3"
Option Explicit

Sub 作用儲存格()
    Range("A1").Select
    ActiveSheet.Shapes("Picture 1").Select
    MsgBox "作用儲存格: " & ActiveCell.Address _
    & Chr(10) & "選擇: " & TypeName(Selection)
End Sub


