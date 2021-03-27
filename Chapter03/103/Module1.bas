Attribute VB_Name = "Module1"
Option Explicit

Sub �ǫh�d��()
    Dim xRange As Range, yRange As Range
    
    If ActiveSheet.FilterMode Then
       ActiveSheet.ShowAllData
    Else
       Set xRange = Range("A6").CurrentRegion
       Set yRange = Range("A1").CurrentRegion
       xRange.AdvancedFilter _
           Action:=xlFilterInPlace, CriteriaRange:=yRange
       Set xRange = Nothing: Set yRange = Nothing
    End If
End Sub

Sub �ǫh�d��2()
    Dim xRange As Range, yRange As Range
    Dim sName, allName
    Set xRange = Worksheets(1).Range("A6").CurrentRegion
    allName = Array("�x�_����", "�s�ˤ���", "�x������", "��������")
    For Each sName In allName
        Range("C2").Value = sName
        Set yRange = Worksheets(1).Range("A1").CurrentRegion
        xRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=yRange, _
                              CopyToRange:=Worksheets(sName).Range("A1")
    Next
    Set xRange = Nothing: Set yRange = Nothing
End Sub

