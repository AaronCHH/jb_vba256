VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Cmd1_Click()
    Call 開始抽選
    a = Application.WorksheetFunction.RandBetween(0, 9)
End Sub

Sub 開始抽選()
    Dim t As Single, r As Range
    Randomize
    t = Timer + 2
    Do Until Timer > t
       For Each r In Range("A2:C2")
           r.Value = Int(10 * Rnd)
       Next
    Loop
End Sub


