Attribute VB_Name = "Module1"
Sub �˵��}�C()
    Dim myArray As Variant
    MsgBox IsArray(myArray)
    
    myArray = Array("���p��", #9/10/2009#, "AB")
    MsgBox IsArray(myArray)
End Sub

