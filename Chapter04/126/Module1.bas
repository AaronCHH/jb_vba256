Attribute VB_Name = "Module1"
Option Explicit

Sub 保護工作表確認()
    If ActiveSheet.ProtectContents Then
        MsgBox "工作表處於保護狀態!!"
        Exit Sub
    End If
    Range("C4:D13").Locked = False
    ActiveSheet.Protect Password:="hogo"
End Sub



