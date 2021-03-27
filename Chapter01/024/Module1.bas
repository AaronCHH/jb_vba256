Attribute VB_Name = "Module1"
Option Explicit

Sub 輸入指定資料()
    Dim myAge As Variant
    myAge = Application.InputBox("請輸入年齡", _
           "輸入年齡", Type:=1)
    If TypeName(myAge) = "Boolean" Then
       Range("B4").Value = "非公開"
    Else
       Range("B4").Value = myAge
    End If
End Sub


