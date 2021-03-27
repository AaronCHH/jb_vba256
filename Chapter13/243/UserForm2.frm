VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "§Æ±æÃþ§O"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
      
   Dim i As Integer, str As String
   For i = 0 To ListBox1.ListCount - 1
       If ListBox1.Selected(i) Then
          Cells(Rows.Count, 1).End(xlUp).Offset(1).Value = ListBox1.List(i)
       End If
   Next
   
End Sub

Private Sub TextBox1_Change()
    TextBox2.Text = Application.GetPhonetic(TextBox1.Text)
End Sub

Private Sub UserForm_Initialize()

   ListBox1.AddItem "°ê»y"
   ListBox1.AddItem "¶m§ø"
   ListBox1.AddItem "·nºu"
   ListBox1.AddItem "¥j¨å"
   ListBox1.AddItem "©Ô¤B"
   ListBox1.AddItem "ÂÅ½Õ"
    
End Sub






