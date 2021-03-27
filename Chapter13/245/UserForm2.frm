VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Åª¨ú¼v¹³"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
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
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "ImageFile", "*.jpg; *.jpeg", 1
        .AllowMultiSelect = False
        If .Show = -1 Then
            Image1.Picture = LoadPicture(.SelectedItems(1))
        End If
    End With
End Sub

