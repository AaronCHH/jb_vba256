Attribute VB_Name = "Module1"

Sub 製作樞紐分析表()
    Dim srcRange As Range
    Dim xPCach As PivotCache, xPTbl As PivotTable
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion
    Worksheets.Add After:=Worksheets(Sheets.Count)
    ActiveSheet.Name = "總計"
    
    Set xPCach = ActiveWorkbook.PivotCaches.Add _
            (SourceType:=xlDatabase, SourceData:=srcRange)
    Set xPTbl = xPCach.CreatePivotTable _
            (TableDestination:=Range("A3"), TableName:="Privot1")
    With xPTbl
        .PivotFields("分店").Orientation = xlColumnField
        .PivotFields("商品名").Orientation = xlRowField
        .PivotFields("金額").Orientation = xlDataField
    End With
End Sub




