Attribute VB_Name = "Module1"
Option Explicit

Sub ��J���w���()
    Dim myAge As Variant
    myAge = Application.InputBox("�п�J�~��", _
           "��J�~��", Type:=1)
    If TypeName(myAge) = "Boolean" Then
       Range("B4").Value = "�D���}"
    Else
       Range("B4").Value = myAge
    End If
End Sub


