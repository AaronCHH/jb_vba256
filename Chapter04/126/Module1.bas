Attribute VB_Name = "Module1"
Option Explicit

Sub �O�@�u�@��T�{()
    If ActiveSheet.ProtectContents Then
        MsgBox "�u�@��B��O�@���A!!"
        Exit Sub
    End If
    Range("C4:D13").Locked = False
    ActiveSheet.Protect Password:="hogo"
End Sub



