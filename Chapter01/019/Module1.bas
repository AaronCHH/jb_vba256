Attribute VB_Name = "Module1"
Option Explicit

Sub ¤Gºû°}¦C()
    Dim myArray(3, 2) As Variant
    Dim i  As Integer, j As Integer
    For i = 0 To 3
        For j = 0 To 2
            myArray(i, j) = Cells(i + 3, j + 2).Value
            Debug.Print myArray(i, j)
        Next j
    Next i
End Sub

Sub ¤Gºû°}¦C2()
    Dim myArray As Variant
    Dim i As Integer, j As Integer
    
    myArray = Range("B3:D6").Value
    For i = 1 To 4
        For j = 1 To 3
            Debug.Print myArray(i, j)
        Next j
    Next i
End Sub
