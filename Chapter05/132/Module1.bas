Attribute VB_Name = "Module1"
Option Explicit

Sub 對話方塊選擇活頁簿()
    With Application.FileDialog(msoFileDialogOpen)
        .FilterIndex = 2
        .AllowMultiSelect = True
        If .Show = -1 Then .Execute
    End With
End Sub

