Attribute VB_Name = "Module1"
Option Explicit

Sub ¦r¦ê¤ñ¸û()
    Dim i As Integer, m1 As String, m2 As String
    For i = 2 To 4
        m1 = Cells(i, 1)
        m2 = Cells(i, 2)
        Cells(i, 3) = IIf(StrComp(m1, m2, vbTextCompare) = 0, "¡³", "¢®")
        Cells(i, 4) = IIf(StrComp(m1, m2, vbBinaryCompare) = 0, "¡³", "¢®")
    Next
End Sub


