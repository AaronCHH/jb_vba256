Attribute VB_Name = "Module1"
Option Explicit

Sub �إߥؿ�()
    Dim myFSO As New FileSystemObject
    Dim myFolder As String
    myFolder = "C:\ExcelVBA\"
    If myFSO.FolderExists(myFolder & "Temp") Then
       MsgBox "�ؿ��s�b"
    Else
       myFSO.CreateFolder (myFolder & "Temp")
    End If
End Sub

Sub �R���ؿ�()
    Dim myFSO As New FileSystemObject
    On Error Resume Next
    myFSO.DeleteFolder "C:\ExcelVBA\Temp"
End Sub

Sub �ƻs�ؿ�()
    Dim myFSO As New FileSystemObject
    On Error Resume Next
    myFSO.CopyFolder "C:\Temp", "C:\Temp2", False
End Sub



