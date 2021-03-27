Attribute VB_Name = "Module1"
Option Explicit

Sub 資料型態轉換()
    Dim myInt As String
    Dim myDate As String
    myInt = "42.195"
    myDate = "民國98年10月10日"
    MsgBox "變換為整數：" & myInt & " → " & CInt(myInt) & Chr(10) & _
           "變換為日期：" & myDate & " → " & CDate(myDate)

End Sub

Sub CInt函數測試()
    Dim myInt As String
    On Error GoTo errMsg
    myInt = "1.5"
    Debug.Print myInt & " → " & CInt(myInt)
    myInt = "2.5"
    Debug.Print myInt & " → " & CInt(myInt)
    myInt = "40000"
    Debug.Print myInt & " → " & CInt(myInt)
    Exit Sub
errMsg:
    MsgBox Err.Number & "：" & Err.Description
End Sub




