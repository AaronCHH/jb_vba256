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
   
   With ListBox1
       If .ListIndex = -1 Then
           MsgBox "請選擇職業!!"
       Else
           Cells(2, 9).Value = .List(.ListIndex)
       End If
   End With
   
End Sub

Private Sub TextBox1_Change()
    TextBox2.Text = Application.GetPhonetic(TextBox1.Text)
End Sub

Private Sub UserForm_Initialize()

   Label1.Caption = 1
   ListBox1.AddItem "打工"
   ListBox1.AddItem "上班族"
   ListBox1.AddItem "自營業"
   ListBox1.AddItem "主婦"
   ListBox1.AddItem "其他"

End Sub





