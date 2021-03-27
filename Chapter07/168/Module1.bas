Attribute VB_Name = "Module1"
Option Explicit

Sub 欄列捲軸()
    Dim myRow As Integer, myCol As Integer
    myRow = Range("商品買家").Row
    myCol = Range("商品買家").Column
    ActiveWindow.ScrollRow = myRow
    ActiveWindow.ScrollColumn = myCol
End Sub

