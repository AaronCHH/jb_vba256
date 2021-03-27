VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "會員登錄"
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
   '文字方塊非數值時的處理
    If Not IsNumeric(kaisu) Then
       MsgBox "請輸入數值!!!"
       TextBox1.Value = SpinButton2.Value
   '文字方塊的值超過微調按鈕時的處理
    ElseIf kaisu < SpinButton2.Min Or kaisu > SpinButton2.Max Then
       MsgBox " 請輸入範圍 " & SpinButton2.Min & "~" & SpinButton2.Max & " 的數值!!!"
       TextBox1.Value = SpinButton2.Value
    Else
   '文字方塊的值設定為微調按鈕的值
       SpinButton2.Value = kaisu
    End If
End Sub

Private Sub UserForm_Click()

End Sub
