Attribute VB_Name = "Module1"
Option Explicit

Sub 轉換目前目錄()
    Range("A2").Value = CurDir
    ChDrive "D"
    ChDir "D:\Work"
    Range("A4").Value = CurDir
End Sub



