Attribute VB_Name = "Module1"
Option Explicit

Sub 重複字串()
    Dim i As Integer
    For i = 2 To 9
        Cells(i, 3).Value = _
        Cells(i, 1) & String(5 - Len(Cells(i, 2)), "0") & Cells(i, 2)
    Next
End Sub

Sub 重複空白()
    Dim Text1 As String, Text2 As String
    Text1 = "建議"
    Text2 = "壽司"
    MsgBox Text1 & Space(3) & Text2
End Sub





