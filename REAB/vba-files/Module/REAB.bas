Attribute VB_Name = "REAB"

Sub Importar()

    'LIMPAR DADOS
    Windows("REAB.xlsm").Activate
    Sheets("DADOS").Select
    Rows("1:105").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=x1Down

    'IMPORTAR DADOS PARA A MACRO
    Windows("loja.xlsx").Activate
    Cells.Select
    Selection.Copy
    Windows("REAB.xlsm").Activate
    Sheets("DADOS").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("MENU").Select
        
End Sub

Sub emergencial()

    ' CONFIRMA��O
    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Voc� est� solicitando a impress�o de REAB EMERGENCIAL. Deseja continuar?", vbYesNo)
        
    If confirmacao = vbYes Then
         
        Sheets("DADOS").Select
        Range("A1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=5, Criteria1:="ZUB"
        
        Columns("B:B").Select
        Selection.Copy
        Sheets("INFORMA��ES").Select
        Range("F1").Select
        ActiveSheet.Paste
        ActiveSheet.Range("$F$1:$F$10000").RemoveDuplicates Columns:=1, Header:= _
            xlYes
            
        lj_total = Range("I2")
        lj = 1
        
        Do While lj < lj_total + 1
            
            ' SELECIONAR LOJA
            lj = lj + 1
            Sheets("INFORMA��ES").Select
            Range("F" & lj).Select
            Selection.Copy
            Sheets("CAPA").Select
            Range("F2").Select
            ActiveSheet.Paste
            
            'MONTAR CAPA
            filtro = Range("F2")
            
            'APARGAR RESIDUOS
            Range("A8:B28").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            
            Sheets("DADOS").Select
            ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=2, Criteria1:=filtro
            ultima = Range("D2").End(xlDown).Row
            Range("D2:D" & ultima).Select
            Selection.Copy
            Sheets("CAPA").Select
            Range("A8").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            'TRANSFORMAR LOJA EM CODIGO EMERGENCIAL
            Range("K2").Select
            Selection.Copy
            Range("F2").Select
            ActiveSheet.Paste
                
            'IMPRIMIR
            Sheets("CAPA").Select
            Range("A1:I36").Select
            Selection.PrintOut Copies:=3, Collate:=True
            
        Loop
        
        Sheets("DADOS").Select
        Selection.AutoFilter
        
        Sheets("MENU").Select
    
    Else
          MsgBox "Voc� cancelou a impress�o!"
    End If

End Sub

Sub chaparia()

    ' CONFIRMA��O
    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Voc� est� solicitando a impress�o de REAB CHAPARIA. Deseja continuar?", vbYesNo)
        
    If confirmacao = vbYes Then

        Sheets("DADOS").Select
        Range("A1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=5, Criteria1:="UB"
        ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=11, Criteria1:="ZCHP"
        
        Columns("B:B").Select
        Selection.Copy
        Sheets("INFORMA��ES").Select
        Range("F1").Select
        ActiveSheet.Paste
        ActiveSheet.Range("$F$1:$F$10000").RemoveDuplicates Columns:=1, Header:= _
            xlYes
            
        lj_total = Range("I2")
        lj = 1
        
        Do While lj < lj_total + 1
            
            ' SELECIONAR LOJA
            lj = lj + 1
            Sheets("INFORMA��ES").Select
            Range("F" & lj).Select
            Selection.Copy
            Sheets("CAPA").Select
            Range("F2").Select
            ActiveSheet.Paste
            
            'MONTAR CAPA
            filtro = Range("F2")
            
            'APARGAR RESIDUOS
            Range("A8:B28").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            
            Sheets("DADOS").Select
            ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=2, Criteria1:=filtro
            ultima = Range("D2").End(xlDown).Row
            Range("D2:D" & ultima).Select
            Selection.Copy
            Sheets("CAPA").Select
            Range("A8").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            'IMPRIMIR
            Sheets("CAPA").Select
            Range("A1:I36").Select
            Selection.PrintOut Copies:=3, Collate:=True
            
        Loop
        
        Sheets("DADOS").Select
        Selection.AutoFilter
        
        Sheets("MENU").Select
    
    Else
          MsgBox "Voc� cancelou a impress�o!"
    End If

End Sub

Sub controle()

    Sheets("DADOS").Select
    Range("A1").Select
    'Selection.AutoFilter
    'ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=5, Criteria1:="UB"
    'ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=11, Criteria1:="ZCHP"
    
    Columns("B:B").Select
    Selection.Copy
    Sheets("INFORMA��ES").Select
    Range("F1").Select
    ActiveSheet.Paste
    ActiveSheet.Range("$F$1:$F$10000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
    lj_total = Range("I2")
    lj = 1
    controle_linha = 2
    
    Do While lj < lj_total + 1
        
        ' SELECIONAR LOJA
        lj = lj + 1
        Sheets("INFORMA��ES").Select
        Range("F" & lj).Select
        numero_loja = Range("F" & lj).Value
        Selection.Copy
        Sheets("CONTROLE").Select
        Range("B" & controle_linha).Select
        ActiveSheet.Paste
        
        'filtrar a loja
        Sheets("DADOS").Select
        Range("A1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$K$10000").AutoFilter Field:=2, Criteria1:=numero_loja
        
        'montar dados
        ultima = Range("D2").End(xlDown).Row
        Range("D2:D" & ultima).Select
        Selection.Copy
        Sheets("CONTROLE").Select
        Range("A" & controle_linha).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("C" & controle_linha).Select
        ActiveCell.FormulaR1C1 = _
            "=CONCATENATE(RC[-2],R[1]C[-2],R[2]C[-2],R[3]C[-2],R[4]C[-2],R[5]C[-2],R[6]C[-2],R[7]C[-2],R[8]C[-2],R[9]C[-2],R[10]C[-2],R[11]C[-2],R[12]C[-2],R[13]C[-2],R[14]C[-2],R[15]C[-2],R[16]C[-2],R[17]C[-2],R[18]C[-2],R[19]C[-2],R[20]C[-2],R[21]C[-2],R[22]C[-2],R[23]C[-2],R[24]C[-2],R[25]C[-2],R[26]C[-2],R[27]C[-2],R[28]C[-2])"
        Range("D" & controle_linha).Select
        ActiveCell.FormulaR1C1 = "=SUBSTITUTE(RC[-1],""4500"",""/"")"
        Range("D" & controle_linha).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        'Peso
        Sheets("DADOS").Select
        Range("F2:F" & ultima).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("CONTROLE").Select
        Range("E" & controle_linha).Select
        ActiveSheet.Paste
        Range("F" & controle_linha).Select
        ActiveCell.FormulaR1C1 = _
            "=SUM(RC[-1],R[1]C[-1],R[2]C[-1],R[3]C[-1],R[4]C[-1],R[5]C[-1],R[6]C[-1],R[7]C[-1],R[8]C[-1],R[9]C[-1],R[10]C[-1],R[11]C[-1],R[12]C[-1],R[13]C[-1],R[14]C[-1],R[15]C[-1],R[16]C[-1],R[17]C[-1],R[18]C[-1],R[19]C[-1],R[20]C[-1],R[21]C[-1],R[22]C[-1],R[23]C[-1],R[24]C[-1],R[25]C[-1],R[26]C[-1],R[27]C[-1],R[28]C[-1])"
        Range("F" & controle_linha).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        
        Range("C" & controle_linha).Select
        Selection.ClearContents
        Range("A:A").Select
        Selection.ClearContents
        Range("E:E").Select
        Selection.ClearContents
        
        controle_linha = controle_linha + 1
            
    Loop
    
    'preencher tp
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],INFORMA��ES!C[-2]:C[1],4,0)"
    Selection.AutoFill Destination:=Range("C2:C100"), Type:=xlFillDefault

    'preencher datas
    Sheets("MENU").Select
    Range("L18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("CONTROLE").Select
    Range("G2:G100").Select
    ActiveSheet.Paste
    Sheets("MENU").Select
    Range("L19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("CONTROLE").Select
    Range("H2:H100").Select
    ActiveSheet.Paste
    
    'reset
    lj_total = lj_total + 2
    
    Rows(lj_total & ":209").Select
    Selection.Delete Shift:=xlUp
    Sheets("DADOS").Select
    Selection.AutoFilter
    
    Sheets("MENU").Select
    
End Sub

