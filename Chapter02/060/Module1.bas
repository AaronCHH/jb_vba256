Attribute VB_Name = "Module1"
Option Explicit

Sub ���o�C�ƩM���()
    Dim rcnt As Long, ccnt As Long
    
    rcnt = Rows.Count
    ccnt = Columns.Count
    MsgBox "�u�@���C��: " & rcnt & Chr(10) & _
           "�u�@�����: " & ccnt
    rcnt = Range("A3:C10").Rows.Count
    ccnt = Range("A3:C10").Columns.Count
    MsgBox "���C��: " & rcnt & Chr(10) & _
           "�����: " & ccnt

End Sub

