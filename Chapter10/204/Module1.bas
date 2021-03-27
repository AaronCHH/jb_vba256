Attribute VB_Name = "Module1"
Option Explicit

Sub 轉換格式()
    Debug.Print "日期: " & Format(Now, "Long Date")
    Debug.Print "時間: " & Format(Now, "hh時nn分")
    Debug.Print "數值: " & Format(25000, "Standard")
    Debug.Print "字串: " & Format("strawberry", ">")
End Sub

