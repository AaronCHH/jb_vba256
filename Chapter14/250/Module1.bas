Attribute VB_Name = "Module1"
Option Explicit

Sub 操作應用程式()
    Dim taskID As Double
    Dim kazu1 As Long, kazu2 As Long
    kazu1 = Range("A2").Value
    kazu2 = Range("C2").Value
    taskID = Shell("calc.exe", vbNormalFocus)
    SendKeys kazu1 & "{*}" & kazu2 & "{ENTER}", True
    SendKeys "^C", True
    Application.Wait Now + TimeValue("0:00:01")
    Range("E2").PasteSpecial xlPasteAll
End Sub


Sub 啟動記事本()
    Shell "notepad.exe ""textdata.txt""", vbNormalFocus
End Sub

