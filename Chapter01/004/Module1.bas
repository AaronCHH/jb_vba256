Attribute VB_Name = "Module1"
Sub HensuTest()

    Dim myName As String
    Dim myDate As Date, myAge As Integer
    Dim myHeight As Single
    
    myName = "王小花"
    myDate = #4/8/1997#
    myAge = Range("A1").Value
    myHeight = 142.3

    MsgBox "姓 名:" & myName & Chr(10) & "生 日:" & myBirth & _
           "年 齡:" & myAge & Chr(10) & "身 高:" & myHeight
           
End Sub
