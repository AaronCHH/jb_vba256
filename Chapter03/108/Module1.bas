Attribute VB_Name = "Module1"
Option Explicit

Sub �������()
    Range("A5").Group _
        Start:=True, End:=True, Periods:=Array(False, False, False, False, _
        True, False, False)
End Sub

Sub �s�@�ϯä��R��()
    Dim srcRange As Range
    Dim xPCach As PivotCache, xPTbl As PivotTable
    
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion
    Worksheets.Add After:=Worksheets(Sheets.Count)
    ActiveSheet.Name = "���p"
    Set xPCach = ActiveWorkbook.PivotCaches.Add _
            (SourceType:=xlDatabase, SourceData:=srcRange)
    
    Set xPTbl = xPCach.CreatePivotTable _
            (TableDestination:=Range("A3"), TableName:="Pivot1")
    
    With xPTbl
        .PivotFields("�ӫ~�W").Orientation = xlColumnField
        .PivotFields("���").Orientation = xlRowField
        .PivotFields("���B").Orientation = xlDataField
    End With
End Sub
