Attribute VB_Name = "Module1"
Option Explicit

Sub �O�@����ï()
    ActiveWorkbook.Protect Password:="PassWord", _
                           Structure:=True, _
                           Windows:=True
End Sub

Sub �Ѱ�����ï()
    ActiveWorkbook.Unprotect Password:="PassWord"
End Sub

