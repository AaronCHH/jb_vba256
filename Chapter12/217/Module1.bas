Attribute VB_Name = "Module1"
Option Explicit

Sub �Ϫ�d���ܧ�()
    Dim sd As Variant, gRange As Range
    Charts("�Ϫ�").Activate
    sd = InputBox("���w�Ϫ�ƪ����:  ��y: 1, �^�y 2, �ƾ� 3")
    Select Case sd
        Case 1: Set gRange = Worksheets("3�Ь�").Range("B3:D13")
        Case 2: Set gRange = Worksheets("3�Ь�").Range("B19:D29")
        Case 3: Set gRange = Worksheets("3�Ь�").Range("B35:E45")
        Case Else
            MsgBox "���w�����T!!"
            Exit Sub
    End Select
    Charts("�Ϫ�").SetSourceData gRange
    Set gRange = Nothing
End Sub




