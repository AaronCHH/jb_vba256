Attribute VB_Name = "Module1"
Option Explicit

Sub HensuTest()

    Dim myName As String
    Dim myDate As Date, myAge As Integer
    Dim myHeight As Single
    
    myName = "���p��"
    myBirth = #4/8/1997#
    myAge = Range("A1").Value
    myHeight = 142.3
    
    MsgBox "�m �W:" & myName & Chr(10) & "�� ��:" & myBirth & _
           "�~ ��:" & myAge & Chr(10) & "�� ��:" & myHeight
    

End Sub

