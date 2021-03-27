Attribute VB_Name = "Module1"
Option Explicit

Sub 移動檔案()
    Dim myFSO As New FileSystemObject
    Dim myFile As String, myFolder As String
    On Error GoTo errHandler
    myFile = "C:\ExcelVBA\Temp\05*.xls"
    myFolder = "C:\ExcelVBA\05月"
    myFSO.MoveFile myFile, myFolder
    Exit Sub
errHandler:
   MsgBox Err.Number & ":" & Err.Description
End Sub

Sub 檔案的複製和刪除()
    Dim myFSO As New FileSystemObject
    Dim myFile As String, myFolder As String

    myFile = "C:\ExcelVBA\Temp\05*.xls"
    myFolder = "C:\ExcelVBA\05月"
    myFSO.CopyFile myFile, myFolder, True
    myFSO.DeleteFile myFile
End Sub








