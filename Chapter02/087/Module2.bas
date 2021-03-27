Attribute VB_Name = "Module2"
Option Explicit

Sub 取得值()
    
    MsgBox "使用Value屬性取得: " & Range("A5").Value & Chr(10) & _
           "使用Text屬性取得: " & Range("A5").Text
End Sub


