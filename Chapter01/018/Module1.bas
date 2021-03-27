Attribute VB_Name = "Module1"
Option Explicit

Sub 陣列的下限值與上限值()
    Dim myArray1(1 To 3) As Integer, myArray2 As Variant
    myArray2 = Array("陳大同", #10/10/2009#, "A", "台北市", "未婚")
    
    MsgBox "下限值" & LBound(myArray1) & _
           "﹑上限值" & UBound(myArray1), , "myarray1的下限值‧上限值"

    MsgBox "下限值" & LBound(myArray2) & _
           "﹑上限值" & UBound(myArray2), , "myarray2的下限值‧上限值"
           
End Sub


