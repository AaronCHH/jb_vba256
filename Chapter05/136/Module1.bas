Attribute VB_Name = "Module1"
Option Explicit

Sub �ˬd�O�_�������A�s��()
    Dim fPath As String, fName As String
    fPath = "C:\ExcelVBA\"
    fName = Format(Date, "mm_dd") & "���G"
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


