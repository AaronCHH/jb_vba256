Attribute VB_Name = "Module1"
Option Explicit

Sub ���~�B�z()
  On Error GoTo errHandler
  Dim bName As String
  
  bName = "C:\ ExcelVBA\Book1.xls"
  Workbooks.Open bName
  
  Exit Sub
errHandler:
  MsgBox "�䤣���ɮ�" & Chr(10) & _
          "�ɮצW��:" & bName
End Sub

Sub ���~�B�z2()
    On Error GoTo errHandler
    Dim bName As String
    
    bName = "C:\ExcelVBA\Book1.xls"
    Workbooks.Open bName
    
    Exit Sub
errHandler:
        MsgBox "�䤣���ɮ�" & Chr(10) & _
        Err.Number & " : " & Err.Description
End Sub


