Attribute VB_Name = "Module1"
Option Explicit

Sub ����ï����s�s�ɽT�{()
    If ActiveWorkbook.Saved Then
        MsgBox "����ï���ݦs��!!!"
    Else
        MsgBox "�w�ק�!!��s�s��!!"
        ActiveWorkbook.Save
    End If
End Sub

Sub �Ҧ�������ï���s�ɪ�������()
    Dim xBook As Workbook
    For Each xBook In Workbooks
        xBook.Saved = True
        xBook.Close
    Next
End Sub


