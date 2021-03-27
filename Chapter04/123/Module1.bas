Attribute VB_Name = "Module1"
Option Explicit

Sub 保護工作表()
    On Error Resume Next
    Range("C4:D13").Locked = False
    ActiveSheet.Protect Password:="hogo"
End Sub

Sub 保護工作表解除()
    On Error GoTo errHander
    ActiveSheet.Unprotect
    Exit Sub
errHander:
    MsgBox "請輸入正確密碼!!!"
End Sub




