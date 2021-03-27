Attribute VB_Name = "Module1"
Option Explicit

Sub 活頁簿的存檔地點()
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

Sub 活頁簿的存檔地點2()
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


