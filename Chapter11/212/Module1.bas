Attribute VB_Name = "Module1"
Option Explicit

Function SEIKYU(KINGAKU As Long) As Long
    Application.Volatile
    Select Case KINGAKU
        Case Is >= 100000
            SEIKYU = KINGAKU * (1 - 0.08)
        Case Is >= 50000
            SEIKYU = KINGAKU * (1 - 0.05)
        Case Else
            SEIKYU = KINGAKU
    End Select
End Function


