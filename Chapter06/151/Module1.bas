Attribute VB_Name = "Module1"
Option Explicit

Sub �R���ؿ�()
    Dim xFolder As String, ans As Integer
    xFolder = "C:\ExcelVBA\Temp"
    If Dir(xFolder, vbDirectory) = "" Then
       MsgBox "�ؿ����s�b!!"
    ElseIf Dir(xFolder & "\*.*", vbNormal) = "" Then
       RmDir xFolder
    Else
       ans = MsgBox("�b[" & xFolder & "]�����ɮצs�b!!" & _
           Chr(10) & "�i�H�R����??", vbYesNo)
       If ans = vbYes Then
          Kill xFolder & "\*.*"
          If Dir(xFolder & "\*.*", vbNormal) = "" Then
             RmDir xFolder
          End If
       End If
    End If
End Sub






