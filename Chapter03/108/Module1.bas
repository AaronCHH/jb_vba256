Attribute VB_Name = "Module1"
Option Explicit

Sub 日期分組()
    Range("A5").Group _
        Start:=True, End:=True, Periods:=Array(False, False, False, False, _
        True, False, False)
End Sub

Sub 製作樞紐分析表()
    Dim srcRange As Range
    Dim xPCach As PivotCache, xPTbl As PivotTable
    
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion
    Worksheets.Add After:=Worksheets(Sheets.Count)
    ActiveSheet.Name = "集計"
    Set xPCach = ActiveWorkbook.PivotCaches.Add _
            (SourceType:=xlDatabase, SourceData:=srcRange)
    
    Set xPTbl = xPCach.CreatePivotTable _
            (TableDestination:=Range("A3"), TableName:="Pivot1")
    
    With xPTbl
        .PivotFields("商品名").Orientation = xlColumnField
        .PivotFields("日期").Orientation = xlRowField
        .PivotFields("金額").Orientation = xlDataField
    End With
End Sub
