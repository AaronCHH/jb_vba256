Attribute VB_Name = "Module1"
Option Explicit

Sub 數值確認()
    Dim tokuten As String
    tokuten = InputBox("請輸入得分!!")
    If IsNumeric(tokuten) Then
       MsgBox tokuten & "分!!"
    Else
       MsgBox "請輸入數值!!"
    End If
End Sub


