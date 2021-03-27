Attribute VB_Name = "Module1"
Option Explicit

Sub 計算式()
    Range("D4").Formula = "=B4*C4"
    Range("D5:D6").FormulaR1C1 = Range("D4").FormulaR1C1

    Range("C7").Formula = "=SUM(C4:C6)"
    Range("D7").FormulaR1C1 = Range("C7").FormulaR1C1
End Sub

Sub 計算式2()
    Range("D4:D6").FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("C7:D7").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
End Sub

