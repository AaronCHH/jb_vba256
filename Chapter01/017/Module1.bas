Attribute VB_Name = "Module1"
Option Explicit

Sub �ܧ󤸯��Ӽ�1()
    Dim myArray() As String, i As Integer
    ReDim myArray(1)
    myArray(0) = "���j��"
    myArray(1) = "�i�p��"
    ReDim myArray(2)
    myArray(2) = "���j�i"
    For i = 0 To UBound(myArray)
        Cells(i + 1, 1).Value = myArray(i)
    Next i
End Sub

Sub �ܧ󤸯��Ӽ�2()
    Dim myArray() As String, i As Integer
    ReDim myArray(1)
    myArray(0) = "���j��"
    myArray(1) = "�i�p��"
    ReDim Preserve myArray(2)
    myArray(2) = "���j�i"
    For i = 0 To UBound(myArray)
        Cells(i + 1, 1).Value = myArray(i)
    Next i
End Sub


