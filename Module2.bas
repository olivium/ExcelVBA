Attribute VB_Name = "Module2"
Sub MakeSelection()
Attribute MakeSelection.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MakeSelection Macro
'

'
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
End Sub
