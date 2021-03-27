Attribute VB_Name = "Module1"
Option Explicit

Sub 新增超連結()
    ActiveSheet.Hyperlinks.Add anchor:=Range("B3"), _
    Address:="c:\月間銷售\1月.xls", SubAddress:="1月!A1", TextToDisplay:="1月"
End Sub

Sub 刪除超連結()
    Dim myHyperLink As Hyperlink
    
    For Each myHyperLink In ActiveSheet.Hyperlinks
        myHyperLink.Delete
    Next
End Sub

Sub 新增超連結2()
    ActiveSheet.Hyperlinks.Add anchor:=Range("D3"), _
        Address:="http://www.flag.com.tw/", _
        TextToDisplay:="旗標資訊"
End Sub
Sub 新增超連結3()
    ActiveSheet.Hyperlinks.Add anchor:=Range("B3"), _
        Address:="", _
        SubAddress:="Sheet2!A1", _
        TextToDisplay:=Worksheets(2).Name
End Sub


