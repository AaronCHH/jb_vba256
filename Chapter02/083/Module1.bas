Attribute VB_Name = "Module1"
Option Explicit

Sub �s�W�W�s��()
    ActiveSheet.Hyperlinks.Add anchor:=Range("B3"), _
    Address:="c:\�붡�P��\1��.xls", SubAddress:="1��!A1", TextToDisplay:="1��"
End Sub

Sub �R���W�s��()
    Dim myHyperLink As Hyperlink
    
    For Each myHyperLink In ActiveSheet.Hyperlinks
        myHyperLink.Delete
    Next
End Sub

Sub �s�W�W�s��2()
    ActiveSheet.Hyperlinks.Add anchor:=Range("D3"), _
        Address:="http://www.flag.com.tw/", _
        TextToDisplay:="�X�и�T"
End Sub
Sub �s�W�W�s��3()
    ActiveSheet.Hyperlinks.Add anchor:=Range("B3"), _
        Address:="", _
        SubAddress:="Sheet2!A1", _
        TextToDisplay:=Worksheets(2).Name
End Sub


