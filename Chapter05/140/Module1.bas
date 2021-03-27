Attribute VB_Name = "Module1"
Option Explicit

Sub 保護活頁簿()
    ActiveWorkbook.Protect Password:="PassWord", _
                           Structure:=True, _
                           Windows:=True
End Sub

Sub 解除活頁簿()
    ActiveWorkbook.Unprotect Password:="PassWord"
End Sub

