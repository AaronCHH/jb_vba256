Attribute VB_Name = "Module1"
Option Explicit

Sub �d�ߦP�W�ɮ�()
    Dim myPath As String, myFile As String
    ChDir "C:\ExcelVBA\"
    myFile = Format(Date, "mm_dd") & "���G.xls"
    If Dir(myFile) = "" Then
       Workbooks.Open Filename:="���յ��G���.xls"
       ActiveWorkbook.SaveAs Filename:=myFile
    Else
       Workbooks.Open Filename:=myFile
    End If
End Sub



