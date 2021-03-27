Attribute VB_Name = "Module1"
Option Explicit

Sub 處理後再開始()
  On Error GoTo errHandler
  Dim bName As String, xFD As FileDialog
  bName = "C:\ ExcelVBA\Book1.xls"
  Workbooks.Open bName
  Exit Sub
errHandler:
    Set xFD = Application.FileDialog(msoFileDialogOpen)
    If xFD.Show = 0 Then Exit Sub
    xFD.Execute
    Set xFD = Nothing
    Resume Next
End Sub

