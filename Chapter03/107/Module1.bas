Attribute VB_Name = "Module1"

Sub �s�@�ϯä��R��()
    Dim srcRange As Range
    Dim xPCach As PivotCache, xPTbl As PivotTable
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion
    Worksheets.Add After:=Worksheets(Sheets.Count)
    ActiveSheet.Name = "�`�p"
    
    Set xPCach = ActiveWorkbook.PivotCaches.Add _
            (SourceType:=xlDatabase, SourceData:=srcRange)
    Set xPTbl = xPCach.CreatePivotTable _
            (TableDestination:=Range("A3"), TableName:="Privot1")
    With xPTbl
        .PivotFields("����").Orientation = xlColumnField
        .PivotFields("�ӫ~�W").Orientation = xlRowField
        .PivotFields("���B").Orientation = xlDataField
    End With
End Sub




