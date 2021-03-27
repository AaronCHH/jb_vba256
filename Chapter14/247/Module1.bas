Attribute VB_Name = "Module1"
Option Explicit

Sub •¥∂}§Â¶r¿…()
    Dim myFso As New FileSystemObject, myText As TextStream
    Dim i As Integer, j As Integer, rLine As String, myArray() As String
   
    Set myText = myFso.OpenTextFile(Filename:="textdata1.txt", IOMode:=ForReading)
    i = 0
    Do Until myText.AtEndOfStream
        rLine = myText.ReadLine
        myArray = Split(rLine, ",")
        For j = 0 To UBound(myArray)
            Cells(i + 1, j + 1).Value = myArray(j)
        Next j
        i = i + 1
    Loop
    myText.Close
End Sub


