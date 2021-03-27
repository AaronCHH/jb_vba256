Attribute VB_Name = "Module1"
Option Explicit

Sub 跑じ计1()
    Dim myArray() As String, i As Integer
    ReDim myArray(1)
    myArray(0) = "朝"
    myArray(1) = "眎地"
    ReDim myArray(2)
    myArray(2) = ""
    For i = 0 To UBound(myArray)
        Cells(i + 1, 1).Value = myArray(i)
    Next i
End Sub

Sub 跑じ计2()
    Dim myArray() As String, i As Integer
    ReDim myArray(1)
    myArray(0) = "朝"
    myArray(1) = "眎地"
    ReDim Preserve myArray(2)
    myArray(2) = ""
    For i = 0 To UBound(myArray)
        Cells(i + 1, 1).Value = myArray(i)
    Next i
End Sub


