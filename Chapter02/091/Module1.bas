Attribute VB_Name = "Module1"
Option Explicit

Sub ��_�зǦC����e()
    Range("A3").CurrentRegion.UseStandardHeight = True
    Range("A3").CurrentRegion.UseStandardWidth = True
End Sub

Sub ��Ӥu�@���_�зǦC���M��e()
    Rows.RowHeight = ActiveSheet.StandardHeight
    Columns.ColumnWidth = ActiveSheet.StandardWidth
End Sub

