Attribute VB_Name = "Module1"
Option Explicit

Sub Loop3()
    Dim i As Integer, myWeight As Double
    i = 3
    
    Do
       myWeight = Cells(i, "B").Value
       Select Case myWeight
           Case Is > Cells(1, "B").Value
               Cells(i, "C").Value = "N"
           Case Else
               Cells(i, "C").Value = "Y"
       End Select
       i = i + 1
    Loop While myWeight > Cells(1, "B").Value

End Sub

