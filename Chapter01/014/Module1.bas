Attribute VB_Name = "Module1"
Option Explicit

Sub Array���()
    Dim myArray As Variant
    
    myArray = Array("���p��", #3/10/2009#, "A")
    Range("B2") = myArray(0)
    Range("B3") = myArray(1)
    Range("B4") = myArray(2)
End Sub




