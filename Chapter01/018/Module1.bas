Attribute VB_Name = "Module1"
Option Explicit

Sub �}�C���U���ȻP�W����()
    Dim myArray1(1 To 3) As Integer, myArray2 As Variant
    myArray2 = Array("���j�P", #10/10/2009#, "A", "�x�_��", "���B")
    
    MsgBox "�U����" & LBound(myArray1) & _
           "�N�W����" & UBound(myArray1), , "myarray1���U���ȡE�W����"

    MsgBox "�U����" & LBound(myArray2) & _
           "�N�W����" & UBound(myArray2), , "myarray2���U���ȡE�W����"
           
End Sub


