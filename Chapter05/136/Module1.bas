Attribute VB_Name = "Module1"
Option Explicit

Sub 檢查是否有巨集再存檔()
    Dim fPath As String, fName As String
    fPath = "C:\ExcelVBA\"
    fName = Format(Date, "mm_dd") & "結果"
    If ActiveWorkbook.HasVBProject Then
       ActiveWorkbook.SaveAs _
           Filename:=fPath & fName, _
           FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
       ActiveWorkbook.SaveAs _
           Filename:=fPath & fName, _
           FileFormat:=xlOpenXMLWorkbook
    End If
End Sub


