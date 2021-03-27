Attribute VB_Name = "Module2"
Option Explicit

Sub 传︽ち传()
    With Range("C4:C7")
         .WrapText = Not .WrapText
    End With
End Sub

Sub 传︽ち传2()
Attribute 传︽ち传2.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Range("C4:C7").WrapText = True
            
End Sub

Sub 传︽ち传3()
    
    Range("C4:C7").WrapText = False

End Sub

Sub 传︽ち传4()
    
    Range("C4:C7").ShrinkToFit = True
        
End Sub

Sub 传︽ち传5()
    
    Range("C4:C7").ShrinkToFit = False
        
End Sub

Sub だ澄﹃()
    Range("A1").Justify
End Sub

