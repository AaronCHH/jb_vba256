VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�|���n��"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Cells(2, 10).Value = TextBox1.Text
End Sub

Private Sub SpinButton2_Change()
    TextBox1.Value = SpinButton2.Value
End Sub

Private Sub TextBox1_Change()
    Dim kaisu
    kaisu = TextBox1.Value
   '��r����D�ƭȮɪ��B�z
    If Not IsNumeric(kaisu) Then
       MsgBox "�п�J�ƭ�!!!"
       TextBox1.Value = SpinButton2.Value
   '��r������ȶW�L�L�ի��s�ɪ��B�z
    ElseIf kaisu < SpinButton2.Min Or kaisu > SpinButton2.Max Then
       MsgBox " �п�J�d�� " & SpinButton2.Min & "~" & SpinButton2.Max & " ���ƭ�!!!"
       TextBox1.Value = SpinButton2.Value
    Else
   '��r������ȳ]�w���L�ի��s����
       SpinButton2.Value = kaisu
    End If
End Sub

Private Sub UserForm_Click()

End Sub
