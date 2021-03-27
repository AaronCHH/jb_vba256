Attribute VB_Name = "Module1"
Option Explicit

'Sheet1
Sub RetsuBunKatsu()
    Dim myArray() As String
    Dim i As Integer, j As Integer
    For j = 1 To 4
        myArray = Split(Cells(j, 1), ",")
        For i = 0 To UBound(myArray)
            Cells(j, 3 + i).Value = myArray(i)
        Next i
    Next j
End Sub

'Sheet2
Sub GyoBunKatsu()
    Dim myArray() As String
    Dim i As Integer
    
    myArray = Split(Range("A1"), Chr(10))
    For i = 0 To UBound(myArray)
        Cells(i + 3, 1).Value = myArray(i)
    Next i
End Sub

'Sheet3
Sub GyoBunKatsu2()
    
    Range("A1:A4").TextToColumns _
        Destination:=Range("C1:C4"), _
        DataType:=xlDelimited, _
        Comma:=True
End Sub


