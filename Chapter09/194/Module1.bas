Attribute VB_Name = "Module1"
Option Explicit

Sub �զX���ɶ�()
    Dim myTime As Date, myTime2 As Date, myTime3 As Date
    myTime = TimeSerial(Range("A2"), Range("A3"), 0)
    myTime2 = TimeSerial(Range("A2"), Range("A3") - 10, 0)
    myTime3 = TimeSerial(Range("A2") + 1, Range("A3"), 0)
    MsgBox "�w���ɶ�: " & myTime & Chr(10) & Chr(10) & _
           "�Щ� " & myTime2 & " �e��F!! " & Chr(10) & _
           "�W�L " & myTime3 & " �h�L��!! "
End Sub


