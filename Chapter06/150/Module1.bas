Attribute VB_Name = "Module1"
Option Explicit

Sub �إߥؿ�()
    Dim myFolder As String
    myFolder = "C:\ExcelVBA\Temp"
    If Dir(myFolder, vbDirectory) = "" Then
       MkDir myFolder
    Else
       MsgBox "�P�W�ؿ��s�b!!"
    End If
End Sub


