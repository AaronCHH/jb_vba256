Attribute VB_Name = "Module1"
Option Explicit

Sub ��C���b()
    Dim myRow As Integer, myCol As Integer
    myRow = Range("�ӫ~�R�a").Row
    myCol = Range("�ӫ~�R�a").Column
    ActiveWindow.ScrollRow = myRow
    ActiveWindow.ScrollColumn = myCol
End Sub

