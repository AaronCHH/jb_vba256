Attribute VB_Name = "Module1"
Sub 檢視陣列()
    Dim myArray As Variant
    MsgBox IsArray(myArray)
    
    myArray = Array("陳小華", #9/10/2009#, "AB")
    MsgBox IsArray(myArray)
End Sub

