Attribute VB_Name = "Module1"
Option Explicit

Sub 打開以逗號分隔的文字檔()
    Workbooks.OpenText Filename:="textdata.txt", Startrow:=2, _
                       DataType:=xlDelimited, Comma:=True
End Sub

Sub 以指定資料格式打開()
    Workbooks.OpenText Filename:="textdata.txt", Comma:=True, _
        Fieldinfo:=Array(Array(1, 2), Array(2, 1), Array(3, 9), Array(4, 3))
End Sub

