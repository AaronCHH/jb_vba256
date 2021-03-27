Attribute VB_Name = "Module1"
Option Explicit

Sub Array函數()
    Dim myArray As Variant
    
    myArray = Array("陳小明", #3/10/2009#, "A")
    Range("B2") = myArray(0)
    Range("B3") = myArray(1)
    Range("B4") = myArray(2)
End Sub




