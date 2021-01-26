Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("I3").Select

End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveCell.FormulaR1C1 = "=RC[-4] & ""/"" &RC[-3]"
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "=siteMapURL_test & RC[-4] & ""/"" & RC[-3]"
    ActiveCell.FormulaR1C1 = "=siteMapURL_test & ""/"" &  RC[-4] & ""/"" & RC[-3]"
    Range("M3").Select
End Sub
