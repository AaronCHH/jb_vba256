Attribute VB_Name = "Module1"
Option Explicit

Sub 寫入文字檔()
    Dim myFso As New FileSystemObject, myText As TextStream
    Dim i As Integer, j As Integer, wLine As String
    Dim myArray() As Variant
    
    Set myText = myFso.OpenTextFile(Filename:="textdata.txt", _
    IOMode:=ForWriting, Create:=True)
    ReDim myArray(Range("A1").CurrentRegion.Columns.Count - 1)
    For i = 1 To Range("A1").CurrentRegion.Rows.Count
        For j = 0 To UBound(myArray)
            myArray(j) = Cells(i, j + 1).Text
            Next
        wLine = Join(myArray, ",")
        myText.WriteLine wLine
    Next
    myText.Close
End Sub

Sub 寫入文字檔2()
    Dim myFso As New FileSystemObject, myText As TextStream
    Dim i As Integer, j As Integer, wLine As String
    Dim myArray() As Variant
    
    Set myText = myFso.OpenTextFile(Filename:="textdata.txt", _
    IOMode:=ForAppending)
    ReDim myArray(Range("A1").CurrentRegion.Columns.Count - 1)
    For i = 1 To Range("A1").CurrentRegion.Rows.Count
        For j = 0 To UBound(myArray)
            myArray(j) = Cells(i, j + 1).Text
            Next
        wLine = Join(myArray, ",")
        myText.WriteLine wLine
    Next
    myText.Close
End Sub



