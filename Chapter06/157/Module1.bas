Attribute VB_Name = "Module1"
Option Explicit

Sub 使用FSO建立磁碟列表()
    Dim myFSO As New FileSystemObject
    Dim myDrive As Drive, i As String
    i = 2
    For Each myDrive In myFSO.Drives
        Cells(i, 1).Value = myDrive.DriveLetter
        i = i + 1
    Next
End Sub


