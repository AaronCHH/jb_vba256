Attribute VB_Name = "Module1"
Option Explicit

Sub 日期確認()
    Dim hizuke As String
    hizuke = InputBox("請輸入今年的日期!!")
    If IsDate(hizuke) Then
       MsgBox Format(CDate(hizuke), "yyyy年mm月dd日")
    Else
       MsgBox "請輸入正確日期!!"
    End If
End Sub

Sub 日期確認2()
    Dim hizuke As String
    hizuke = InputBox("請輸入今年的日期!!")
    If IsDate(hizuke) Then
       If Year(CDate(hizuke)) = Year(Date) Then
          MsgBox Format(CDate(hizuke), "yyyy年mm月dd日")
       Else
          MsgBox "請輸入今年的日期!!"
       End If
    Else
       MsgBox "請輸入正確日期!!"
    End If
End Sub
