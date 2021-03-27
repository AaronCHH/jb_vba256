Attribute VB_Name = "Module1"
Option Explicit

Sub 建立目錄()
    Dim myFSO As New FileSystemObject
    Dim myFolder As String
    myFolder = "C:\ExcelVBA\"
    If myFSO.FolderExists(myFolder & "Temp") Then
       MsgBox "目錄存在"
    Else
       myFSO.CreateFolder (myFolder & "Temp")
    End If
End Sub

Sub 刪除目錄()
    Dim myFSO As New FileSystemObject
    On Error Resume Next
    myFSO.DeleteFolder "C:\ExcelVBA\Temp"
End Sub

Sub 複製目錄()
    Dim myFSO As New FileSystemObject
    On Error Resume Next
    myFSO.CopyFolder "C:\Temp", "C:\Temp2", False
End Sub



