Attribute VB_Name = "Module1"
Option Explicit

Sub ��ƫ��A�ഫ()
    Dim myInt As String
    Dim myDate As String
    myInt = "42.195"
    myDate = "����98�~10��10��"
    MsgBox "�ܴ�����ơG" & myInt & " �� " & CInt(myInt) & Chr(10) & _
           "�ܴ�������G" & myDate & " �� " & CDate(myDate)

End Sub

Sub CInt��ƴ���()
    Dim myInt As String
    On Error GoTo errMsg
    myInt = "1.5"
    Debug.Print myInt & " �� " & CInt(myInt)
    myInt = "2.5"
    Debug.Print myInt & " �� " & CInt(myInt)
    myInt = "40000"
    Debug.Print myInt & " �� " & CInt(myInt)
    Exit Sub
errMsg:
    MsgBox Err.Number & "�G" & Err.Description
End Sub




