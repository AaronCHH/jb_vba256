Attribute VB_Name = "Module1"
Option Explicit

Sub ���o��}()
    Dim myRange As Range
    Set myRange = Range("A3").CurrentRegion
    MsgBox "�x�s��A3���ثe�ϰ�@�λ��: " & myRange.Address(RowAbsolute:=False, _
            ColumnAbsolute:=False, ReferenceStyle:=xlA1, External:=True)
    Set myRange = Nothing
End Sub


