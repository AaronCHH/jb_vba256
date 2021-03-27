Attribute VB_Name = "Module1"
Sub TrandformAllXLSFilesToXLSM()
Dim myPath As String

myPath = ActiveWorkbook.Path + "\"
WorkFile = Dir(myPath & "*.xls")

Do While WorkFile <> ""
    If Right(WorkFile, 4) <> "xlsm" Then
        Workbooks.Open Filename:=myPath & WorkFile
        ActiveWorkbook.SaveAs Filename:= _
        myPath & WorkFile & "m", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
     End If
     WorkFile = Dir()
Loop
End Sub


