Attribute VB_Name = "Module1"
Option Explicit

Sub µ²¦X°}¦C()
    Dim myArray1(1) As String, myArray2(2) As String
    Dim i As Integer
    Worksheets("Sheet1").Select
    For i = 2 To 6
        myArray1(0) = Cells(i, 2).Value
        myArray1(1) = Cells(i, 3).Value
        myArray2(0) = Cells(i, 4).Value
        myArray2(1) = Cells(i, 5).Value
        myArray2(2) = Cells(i, 6).Value
        Worksheets("Sheet2").Cells(i, 2).Value = Join(myArray1)
        Worksheets("Sheet2").Cells(i, 3).Value = Join(myArray2, "")
    Next i
End Sub