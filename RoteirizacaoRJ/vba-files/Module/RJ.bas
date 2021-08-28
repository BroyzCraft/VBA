Attribute VB_Name = "RJ"

Sub Reset()

    ' Confirma se realmente deseja resetar a tabela
    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Deseja apagar todos os registros?", vbYesNo)
    
    If confirmacao = vbYes Then
        
        'dados da planilha de apoio
        Range("B2:B8").Select
        Selection.ClearContents
        Range("J2:K2").Select
        Selection.ClearContents
        Range("J4:K5").Select
        Selection.ClearContents
        Range("M3").Select
        Selection.ClearContents
        Range("P3:T100").Select
        Selection.ClearContents
        Range("Q2:Q3").Select
        Selection.ClearContents
        Range("O:O").Select
        Selection.Copy
        Range("P:S").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Range("O2:S2").Select
        Selection.Copy
        Range("O3:S100").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        'dados da planilha de capas
        Sheets("CAPA").Select
        Range("C14:D40").Select
        Selection.ClearContents
        Range("L14:M40").Select
        Selection.ClearContents
        Range("F14:I40").Select
        Selection.ClearContents
        
    End If
    
    Sheets("APOIO").Select
    
End Sub

Sub OrganizarRotas()

    ActiveWorkbook.Worksheets("APOIO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("APOIO").Sort.SortFields.Add2 Key:=Range("Q7:Q52") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("APOIO").Sort
        .SetRange Range("P7:S52")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Sub ImprimirCortes()

    ' CONFIRMA��O
    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Voc� solicitou a impress�o das capas de corte, Continuar?", vbYesNo)
    
    If confirmacao = vbYes Then
        
        qtd = Application.InputBox("Digite quantas capas deseja imprimir: ")
        Sheets("CAPA").Select
        Range("A1:M44").Select
        Selection.PrintOut Copies:=qtd, Collate:=True
    
    End If
    
    Sheets("APOIO").Select
    
End Sub

Sub ImprimirCapas()

    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Voc� solicitou a impress�o das capas de roteiro, Continuar?", vbYesNo)
    
    Dim imprimir As VbMsgBoxResult
    imprimir = MsgBox("Deseja criar uma nova capa? ", vbYesNo)
    
    Dim nome As String
    nome = Range("B12").Value
    
    If imprimir = vbYes Then
    
        Sheets("BKP (2)").Select
        ActiveSheet.Name = nome
        
    End If
    
    Sheets(nome).Select
    
    If confirmacao = vbYes Then
        
        'Imprimi as capas
        qtd = Application.InputBox("Digite quantas capas deseja imprimir: ")
        Sheets(nome).Select
        Range("A1:J40").Select
        Selection.PrintOut Copies:=qtd, Collate:=True, PrToFileName:="RJ"
              
    End If
    
    confirmacao = MsgBox("Deseja salvar os dados?", vbYesNo)
    
    If confirmacao = vbYes Then
            
        'Consolida os dados
        Sheets(nome).Select
        Range("A1:J40").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A1:J40").Select
        
        'Gerar PDF
        strPathNome = ThisWorkbook.Path & "\" & "RJ - " & nome & ".pdf"
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strPathNome, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    End If
    
End Sub
