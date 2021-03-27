Attribute VB_Name = "Module1"
Option Explicit

Sub 錯誤處理()
  On Error GoTo errHandler
  Dim bName As String
  
  bName = "C:\ ExcelVBA\Book1.xls"
  Workbooks.Open bName
  
  Exit Sub
errHandler:
  MsgBox "找不到檔案" & Chr(10) & _
          "檔案名稱:" & bName
End Sub

Sub 錯誤處理2()
    On Error GoTo errHandler
    Dim bName As String
    
    bName = "C:\ExcelVBA\Book1.xls"
    Workbooks.Open bName
    
    Exit Sub
errHandler:
        MsgBox "找不到檔案" & Chr(10) & _
        Err.Number & " : " & Err.Description
End Sub


