Attribute VB_Name = "Module1"
Option Explicit

Sub �����ɮ�()
    Dim myFSO As New FileSystemObject
    Dim myFile As String, myFolder As String
    On Error GoTo errHandler
    myFile = "C:\ExcelVBA\Temp\05*.xls"
    myFolder = "C:\ExcelVBA\05��"
    myFSO.MoveFile myFile, myFolder
    Exit Sub
errHandler:
   MsgBox Err.Number & ":" & Err.Description
End Sub

Sub �ɮת��ƻs�M�R��()
    Dim myFSO As New FileSystemObject
    Dim myFile As String, myFolder As String

    myFile = "C:\ExcelVBA\Temp\05*.xls"
    myFolder = "C:\ExcelVBA\05��"
    myFSO.CopyFile myFile, myFolder, True
    myFSO.DeleteFile myFile
End Sub








