Attribute VB_Name = "Module1"
Option Explicit

Sub 刪除目錄()
    Dim xFolder As String, ans As Integer
    xFolder = "C:\ExcelVBA\Temp"
    If Dir(xFolder, vbDirectory) = "" Then
       MsgBox "目錄不存在!!"
    ElseIf Dir(xFolder & "\*.*", vbNormal) = "" Then
       RmDir xFolder
    Else
       ans = MsgBox("在[" & xFolder & "]中有檔案存在!!" & _
           Chr(10) & "可以刪除嗎??", vbYesNo)
       If ans = vbYes Then
          Kill xFolder & "\*.*"
          If Dir(xFolder & "\*.*", vbNormal) = "" Then
             RmDir xFolder
          End If
       End If
    End If
End Sub






