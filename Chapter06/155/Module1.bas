Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�ɮ��ݩ�()
    Dim myFile As String
    myFile = "D:\mee\�X�Шư�\�M��\F0034Excel VBA�ت��֬d���ѦҤ�U\�d���ɮ�\Chapter06\ExcelVBA\���յ��G���.xls"
    SetAttr myFile, vbReadOnly
End Sub

Sub ���o�ݩ�()
    On Error Resume Next
    MsgBox "�s�@�H: " & ActiveWorkbook.BuiltinDocumentProperties("Author") & Chr(10) & _
           "�s�@�ɶ�: " & ActiveWorkbook.BuiltinDocumentProperties("Creation date")
End Sub

Sub �R���ݩ�()
    ActiveWorkbook.RemoveDocumentInformation (xlRDIDocumentProperties)
End Sub

