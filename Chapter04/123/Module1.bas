Attribute VB_Name = "Module1"
Option Explicit

Sub �O�@�u�@��()
    On Error Resume Next
    Range("C4:D13").Locked = False
    ActiveSheet.Protect Password:="hogo"
End Sub

Sub �O�@�u�@��Ѱ�()
    On Error GoTo errHander
    ActiveSheet.Unprotect
    Exit Sub
errHander:
    MsgBox "�п�J���T�K�X!!!"
End Sub




