Attribute VB_Name = "Module1"
Option Explicit

Sub 建立目錄()
    Dim myFolder As String
    myFolder = "C:\ExcelVBA\Temp"
    If Dir(myFolder, vbDirectory) = "" Then
       MkDir myFolder
    Else
       MsgBox "同名目錄存在!!"
    End If
End Sub


