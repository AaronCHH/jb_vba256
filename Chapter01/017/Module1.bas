Attribute VB_Name = "Module1"
Option Explicit

Sub 變更元素個數1()
    Dim myArray() As String, i As Integer
    ReDim myArray(1)
    myArray(0) = "陳大明"
    myArray(1) = "張小華"
    ReDim myArray(2)
    myArray(2) = "王大可"
    For i = 0 To UBound(myArray)
        Cells(i + 1, 1).Value = myArray(i)
    Next i
End Sub

Sub 變更元素個數2()
    Dim myArray() As String, i As Integer
    ReDim myArray(1)
    myArray(0) = "陳大明"
    myArray(1) = "張小華"
    ReDim Preserve myArray(2)
    myArray(2) = "王大可"
    For i = 0 To UBound(myArray)
        Cells(i + 1, 1).Value = myArray(i)
    Next i
End Sub


