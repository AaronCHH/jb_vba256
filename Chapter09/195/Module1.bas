Attribute VB_Name = "Module1"
Option Explicit

Sub 回答測定時間()
    Dim s1 As Single, s2 As Double, qText As String, aText As String
    Dim ansText As String, msgText As String
    
    qText = "英國的正式名稱是?? (半形英文字母)"
    aText = "UNITED KINGDOM"
    s1 = Timer
    ansText = InputBox(qText, "問題")
    s2 = Timer - s1
    msgText = "正確答案: " & aText & Chr(10) & "回答時間: " & s2 & "秒"
    If StrComp(aText, ansText, 1) = 0 Then
       MsgBox "正確!!" & Chr(10) & msgText
    Else
       MsgBox "錯誤!!" & Chr(10) & msgText
    End If
End Sub


