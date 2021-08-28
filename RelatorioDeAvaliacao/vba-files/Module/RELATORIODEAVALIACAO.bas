Attribute VB_Name = "RELATORIODEAVALIACAO"
Sub Importar()

    Apagar
    ImportarResumo
    ImportarTp
    Formatar

End Sub
Sub Apagar()

    Sheets("DADOS - RESUMO").Select
    Cells.Select
    Selection.ClearContents
    Sheets("DADOS - SERVICOS").Select
    Cells.Select
    Selection.ClearContents
    Sheets("AUX").Select

End Sub

Sub ImportarResumo()
Attribute ImportarResumo.VB_ProcData.VB_Invoke_Func = " \n14"

    Windows("ava.xlsx").Activate
    Sheets("Resumo").Select
    Cells.Select
    Selection.Copy
    Windows("Relatorio de Avalia��o.xlsm").Activate
    Range("A1").Select
    Sheets("DADOS - RESUMO").Select
    Range("A1").Select
    ActiveSheet.Paste
    
End Sub
Sub ImportarTp()
Attribute ImportarTp.VB_ProcData.VB_Invoke_Func = " \n14"

    Windows("ava.xlsx").Activate
    Sheets("Detalhamento").Select
    Cells.Select
    Selection.Copy
    Windows("Relatorio de Avalia��o.xlsm").Activate
    Sheets("DADOS - SERVICOS").Select
    Range("A1").Select
    ActiveSheet.Paste
    
End Sub
Sub Formatar()
Attribute Formatar.VB_ProcData.VB_Invoke_Func = " \n14"

    quant_linhas = Range("A1").End(xlDown).Row

    Sheets("DADOS - RESUMO").Select
    Columns("P:S").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Style = "Comma"
    
    Columns("P:P").Select
    Selection.TextToColumns Destination:=Range("P1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
        
    Columns("Q:Q").Select
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
        
    Columns("R:R").Select
    Selection.TextToColumns Destination:=Range("R1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
    Columns("S:S").Select
    Selection.TextToColumns Destination:=Range("S1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],AUX!C[-8]:C[-7],2,0)"
    
    Range("I2").Select
    Selection.Copy
    Range("I3:I" & quant_linhas).Select
    ActiveSheet.Paste
    
    Columns("I:I").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("I1").Value = Range("J1").Value
    
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    
    Sheets("DADOS - SERVICOS").Select
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC12,CHOOSE({1,2},'DADOS - RESUMO'!R2C14:R500C14,'DADOS - RESUMO'!R2C9:R500C9,),2,0)"
    
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K" & quant_linhas)
    
    Range("K1").Value = "TP"
    
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Columns("Q:Q").Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Style = "Comma"
        
    Columns("Q:Q").Select
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Sheets("DASHBOARD").Select
    
End Sub

Sub FiltroSP()

    With ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Transportadora")
        .SlicerItems("AUTO CLEAN").Selected = True
        .SlicerItems("GST").Selected = True
        .SlicerItems("KGB").Selected = True
        .SlicerItems("LUMA").Selected = True
        .SlicerItems("MOTOBOY").Selected = True
        .SlicerItems("MTRANS").Selected = True
        .SlicerItems("ND").Selected = True
        .SlicerItems("WC").Selected = True
        .SlicerItems("FAGUNDES").Selected = False
        .SlicerItems("J.L SARAIVA").Selected = False
        .SlicerItems("MARCRIS").Selected = False
        .SlicerItems("R. NUNES").Selected = False
        .SlicerItems("LTL").Selected = False
        .SlicerItems("TESTE").Selected = False
        .SlicerItems("#N/D").Selected = False
    End With
    
    With ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Transportadora2")
        .SlicerItems("AUTO CLEAN").Selected = True
        .SlicerItems("GST").Selected = True
        .SlicerItems("KGB").Selected = True
        .SlicerItems("LUMA").Selected = True
        .SlicerItems("MOTOBOY").Selected = True
        .SlicerItems("MTRANS").Selected = True
        .SlicerItems("ND").Selected = True
        .SlicerItems("WC").Selected = True
        .SlicerItems("FAGUNDES").Selected = False
        .SlicerItems("J.L SARAIVA").Selected = False
        .SlicerItems("MARCRIS").Selected = False
        .SlicerItems("R. NUNES").Selected = False
        .SlicerItems("LTL").Selected = False
        .SlicerItems("TESTE").Selected = False
        .SlicerItems("#N/D").Selected = False
    End With
    
End Sub
Sub FiltroInterior()

    With ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Transportadora")
        .SlicerItems("FAGUNDES").Selected = True
        .SlicerItems("J.L SARAIVA").Selected = True
        .SlicerItems("MARCRIS").Selected = True
        .SlicerItems("R. NUNES").Selected = True
        .SlicerItems("LTL").Selected = True
        .SlicerItems("AUTO CLEAN").Selected = False
        .SlicerItems("GST").Selected = False
        .SlicerItems("KGB").Selected = False
        .SlicerItems("LUMA").Selected = False
        .SlicerItems("MOTOBOY").Selected = False
        .SlicerItems("MTRANS").Selected = False
        .SlicerItems("ND").Selected = False
        .SlicerItems("WC").Selected = False
        .SlicerItems("TESTE").Selected = False
        .SlicerItems("#N/D").Selected = False
    End With
    
    With ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Transportadora2")
        .SlicerItems("FAGUNDES").Selected = True
        .SlicerItems("J.L SARAIVA").Selected = True
        .SlicerItems("MARCRIS").Selected = True
        .SlicerItems("R. NUNES").Selected = True
        .SlicerItems("LTL").Selected = True
        .SlicerItems("AUTO CLEAN").Selected = False
        .SlicerItems("GST").Selected = False
        .SlicerItems("KGB").Selected = False
        .SlicerItems("LUMA").Selected = False
        .SlicerItems("MOTOBOY").Selected = False
        .SlicerItems("MTRANS").Selected = False
        .SlicerItems("ND").Selected = False
        .SlicerItems("WC").Selected = False
        .SlicerItems("TESTE").Selected = False
        .SlicerItems("#N/D").Selected = False
    End With
    
End Sub
