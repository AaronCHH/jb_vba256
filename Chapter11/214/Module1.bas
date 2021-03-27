Attribute VB_Name = "Module1"
Option Explicit

Function JUNBAN(ParamArray TEAM() As Variant) As Variant
    Dim x As Variant, str As String
    If IsMissing(TEAM) Then
        JUNBAN = CVErr(xlErrNA)
        Exit Function
    End If
    For Each x In TEAM
        str = str & x & "¡÷"
    Next
    JUNBAN = str & "END"
End Function

