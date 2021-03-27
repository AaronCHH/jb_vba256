Attribute VB_Name = "Module1"
Option Explicit

Sub 建立目錄列表()
    Dim myFSO As New FileSystemObject
    Dim myFolder As Folder, i As Integer, subFolder As Folders
    Set subFolder = myFSO.GetFolder("C:\ExcelVBA").SubFolders
    i = 2
    For Each myFolder In subFolder
        Cells(i, 1).Value = myFolder.Name
        Cells(i, 2).Value = myFolder.DateCreated
        i = i + 1
    Next
End Sub

Sub 建立檔案列表()
    Dim myFSO As New FileSystemObject, myFiles As Files
    Dim myFile As File, i As Integer
    
    Set myFiles = myFSO.GetFolder("C:\ExcelVBA").Files
    i = 2
    For Each myFile In myFiles
        Cells(i, 1).Value = myFile.Name
        Cells(i, 2).Value = myFile.DateCreated
        i = i + 1
    Next
End Sub



