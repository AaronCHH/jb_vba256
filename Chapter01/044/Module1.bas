Attribute VB_Name = "Module1"
Option Explicit

Sub DebugTest()
    Dim xData As String
    Dim i As Integer
    
    For i = 1 To 12
        xData = Cells(i, 1).Value
        Debug.Print i & " : " & xData
    Next i
End Sub

