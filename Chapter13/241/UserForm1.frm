VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "·|­ûµn¿ý"
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
    Cells(2, 8).Value = ComboBox1.Text
End Sub

Private Sub TextBox1_Change()
    TextBox2.Text = Application.GetPhonetic(TextBox1.Text)
End Sub

Private Sub UserForm_Initialize()
    Label1.Caption = 1
    ComboBox1.RowSource = Range("¿¤¥«").Address
End Sub





