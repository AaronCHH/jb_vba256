Attribute VB_Name = "Module1"
Option Explicit

Sub ����ï���s�ɦa�I()
    Dim fPath As String
    On Error GoTo errHandler
    
    fPath = ActiveWorkbook.Path
    ActiveSheet.Copy
    ActiveWorkbook.SaveAs Filename:=fPath & "\" & Range("A1").Value & "xls"
    MsgBox ActiveWorkbook.Path
    Exit Sub
errHandler:
    MsgBox Err.Description
End Sub

Sub ����ï���s�ɦa�I2()
    On Error GoTo errHandler
     
    ChDir ActiveWorkbook.Path
    ActiveSheet.Copy
    ActiveWorkbook.SaveAs _
        Filename:=Range("A1").Value & ".xls"
    MsgBox ActiveWorkbook.Path
    Exit Sub
errHandler:
    MsgBox Err.Description
End Sub


