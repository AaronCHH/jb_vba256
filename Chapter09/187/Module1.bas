Attribute VB_Name = "Module1"
Option Explicit

Sub 個別取得時分秒()
    MsgBox "現在的時間" & Chr(10) & _
            Hour(Now) & "時" & Chr(10) & _
            Minute(Now) & "分" & Chr(10) & _
            Second(Now) & "秒"
End Sub

Sub 指定時分秒測試()
    MsgBox Hour("8時35分 PM") & Chr(10) & _
            Minute("20:35") & Chr(10) & _
            Second("下午8時35分")
End Sub





