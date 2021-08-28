Attribute VB_Name = "FAROLRE"

Sub Importacao()

    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Deseja atualizar o farol ? (Deixe a planilha extraida do routeasy aberta e renomeada para 'farol')", vbYesNo)
        
    If confirmacao = vbYes Then

        'APAGAR DADOS ANTERIORES
        Sheets("DADOS BRUTOS").Select
        Range("A1").Select
        Selection.AutoFilter
        Cells.Select
        Selection.ClearContents
        
        Sheets("DADOS").Select
        Range("A2:C300").Select
        Selection.ClearContents
        
        'IMPORTAR DADOS
        Windows("farol.xlsx").Activate
        Range("A1:AA300").Select
        Selection.Copy
        Windows("Farol RoutEasy.xlsm").Activate
        Sheets("DADOS BRUTOS").Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        'FILTRAR TRANSPORTADORAS
        Sheets("DADOS").Select
        quant_trans = Range("L2").Value
        aux = 0
        inicio = 3
        
        Do While aux < quant_trans
        
            Sheets("DADOS").Select
            trans = Range("K" & inicio + aux)
            
            Sheets("DADOS BRUTOS").Select
            ActiveSheet.Range("$A$1:$T$106").AutoFilter Field:=3, Criteria1:="=*" & trans & "*", _
            Operator:=xlAnd
            
            Range("I2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("DADOS").Select
            
            quant_atual = Range("I3").Value + 2
            
            Range("A" & quant_atual).Select
            ActiveSheet.Paste
            
            aux = aux + 1
        Loop
    
        quant_atual = Range("I3").Value + 1
        Range("A2:A" & quant_atual).Select
        Selection.TextToColumns Destination:=Range("B2"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

    Else
          MsgBox "Voc� cancelou a impress�o!"
    End If
    
End Sub
