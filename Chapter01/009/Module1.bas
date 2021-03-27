Attribute VB_Name = "Module1"
Option Explicit

Sub TypeNameTest()
    Dim myVar As Object
    Set myVar = ActiveSheet
    
    Select Case TypeName(myVar)
        Case "Worksheet"
             myVar.PrintPreview
        Case "Chart"
             MsgBox "請選擇工作表!!"
    End Select
End Sub

Sub TypeNameTest2()
    Dim myVar2
    myVar2 = Selection.Value
    MsgBox TypeName(myVar2)
End Sub
