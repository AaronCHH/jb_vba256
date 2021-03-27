Attribute VB_Name = "Module1"
Option Explicit

Sub 查詢同名檔案()
    Dim myPath As String, myFile As String
    ChDir "C:\ExcelVBA\"
    myFile = Format(Date, "mm_dd") & "結果.xls"
    If Dir(myFile) = "" Then
       Workbooks.Open Filename:="測試結果表單.xls"
       ActiveWorkbook.SaveAs Filename:=myFile
    Else
       Workbooks.Open Filename:=myFile
    End If
End Sub



