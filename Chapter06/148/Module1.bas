Attribute VB_Name = "Module1"
Option Explicit

Sub 取得目前目錄()
    Dim xFSO As New FileSystemObject
    Dim xDrive, i As Integer
    i = 2
    For Each xDrive In xFSO.Drives
        Cells(i, 1) = xDrive.DriveLetter
        Cells(i, 2) = CurDir(Cells(i, 1))
        i = i + 1
    Next
    Set xFSO = Nothing
End Sub



