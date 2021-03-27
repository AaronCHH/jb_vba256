Attribute VB_Name = "Module1"
Option Explicit

Sub ÀÉ®×¦Cªí()
    Dim myFolder As String, myFile As String
    Dim i As Integer
    myFolder = "C:\ExcelVBA\"
    myFile = Dir(myFolder & "*.xls?", vbNormal)
    i = 3
    Do While myFile <> ""
        Cells(i, 1).Value = myFile
        myFile = Dir()
        i = i + 1
    Loop
End Sub




