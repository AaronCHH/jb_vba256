Attribute VB_Name = "Module1"
Option Explicit

Sub 輸入指令()
    Dim myComment As String
    myComment = InputBox("請鍵入20字以內的說明", _
                "輸入說明", "來自" & Range("A4") & "先生(小姐)的訊息")
    If Len(myComment) > 20 Then
       MsgBox "字串過長。" & Len(myComment)
       Exit Sub
    End If
    Range("B4").Value = myComment
End Sub


