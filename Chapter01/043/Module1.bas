Attribute VB_Name = "Module1"
Option Explicit

Sub GyoBunKatsu()
    Dim myArray() As String
    Dim i As Integer
    
    myArray = Split(Range("A1"), Chr(10))
    For i = 0 To UBound(myArray)
        Cells(i + 3, 1).Value = myArray(i)
    Next i
End Sub
