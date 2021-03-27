Attribute VB_Name = "Module1"
Option Explicit

Sub 陣列練習()
    Dim xArray(2) As String
    xArray(0) = "王小華"
    xArray(1) = "民國98年10月10日"
    xArray(2) = "A"
    Range("B2") = xArray(0)
    Range("B3") = CDate(xArray(1))
    Range("B4") = xArray(2)
End Sub

