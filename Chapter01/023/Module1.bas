Attribute VB_Name = "Module1"
Option Explicit

Sub ��J���O()
    Dim myComment As String
    myComment = InputBox("����J20�r�H��������", _
                "��J����", "�Ӧ�" & Range("A4") & "����(�p�j)���T��")
    If Len(myComment) > 20 Then
       MsgBox "�r��L���C" & Len(myComment)
       Exit Sub
    End If
    Range("B4").Value = myComment
End Sub


