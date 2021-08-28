Attribute VB_Name = "EMAILLOJAS67"

Function rangetoHTML(rng As Range, layout As String)

    Dim fso As Object
    Dim ts As Object
    Dim tempFile As String
    Dim tempWB As Workbook
    
    tempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".html"
    
    rng.Copy
    Set tempWB = ThisWorkbook
    
    With tempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=tempFile, _
         Sheet:=tempWB.Sheets(layout).Name, _
         Source:=tempWB.Sheets(layout).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
         .Publish (True)
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(tempFile).OpenAsTextStream(1, -2)
    rangetoHTML = ts.readALL
    ts.Close
    rangetoHTML = Replace(rangetoHTML, "align=center x:publishsource=", "align=left x:publishsource=")
    
    Kill tempFile
    
    Set ts = Nothing
    Set fso = Nothing
    Set tempWB = Nothing
           
End Function

Sub EnviarEmail()
    
    'Enviar o e-mail para todos os destinat�rios
    
    Application.DisplayAlerts = False 'Desabilitar alertas
    Application.ScreenUpdating = False 'Desabilitar atualiza��o de tela
    
    ActiveWorkbook.Save 'Salvar planilha
    
    Dim outApp As Outlook.Application 'Vari�vel da aplica��o do outlook
    Dim outMail As Outlook.MailItem 'Vari�vel do objeto e-mail
    
    Dim sh_capa, sh_layout As Worksheet 'Vari�veis das abas do excel
    Dim nome_layout As String 'Nome da aba layout
    Dim destino As String 'Destino email
    
    Dim rng As Range 'Vari�vel da fun��o rangetoHTML
    
    nome_layout = "LAYOUT"
    
    Set sh_capa = Sheets("MENU") 'Configura aba Capa
    Set sh_layout = Sheets(nome_layout) 'Configura aba BD

    Dim i, j, k, l, m, num, lin_layout, lin_bd, num_prod, lin_tabela, cont_email As Long 'Vari�veis auxiliares
    
    'Encontra a linha que a tabela da aba layout come�a
    k = 1
    Do While sh_layout.Cells(k, "B").Value <> "LOJA"
        k = k + 1
    Loop
    'Linha inicial da tabela layout
    lin_tabela = k + 1
    
    'Seleciona a aba Capa
    sh_capa.Select
    
    'Se n�o estiver com o filtro ativo, ativa o filtro
    'If Not sh_capa.AutoFilterMode Then
        'sh_capa.Range("B2:I2").AutoFilter
    'End If
    
    'Filta os dados em ordem crescente de nome
    'sh_capa.AutoFilter.Sort.SortFields.Clear
    'sh_capa.AutoFilter.Sort.SortFields.Add Key:=Range( _
    '"C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    'xlSortNormal
    'With sh_capa.AutoFilter.Sort
        '.Header = xlYes
        '.MatchCase = False
        '.Orientation = xlTopToBottom
        '.SortMethod = xlPinYin
        '.Apply
    'End With
    
    'Se estiver com o filtro ativo,desativa o filtro
    'If sh_capa.AutoFilterMode Then
        'sh_capa.Range("B2:I2").AutoFilter
    'End If
    
    On Error Resume Next 'Habilita tratamento de erros
        Set outApp = GetObject(, "OUTLOOK.APPLICATION") 'Tenta configurar a aplica��o do outlook
        If (outApp Is Nothing) Then 'Se outlook n�o estiver aberto...
            Set outApp = CreateObject("OUTLOOK.APPLICATION") 'Configura a aplica��o do outlook
        End If
    On Error GoTo 0 'Desabilita tratamento de erros
    
    'Linha inicial do bando de dados
    lin_bd = 3
    
    'Roda para todas as linhas do banco de dados
    i = 0
    cont_email = 0
    Do While sh_capa.Cells(lin_bd + i, "B").Value <> ""
        
        'Conta quantos produtos cada respons�vel possui
        num_prod = 0
        Do While sh_capa.Cells(lin_bd + i, "B").Value = sh_capa.Cells(lin_bd + i + num_prod, "B").Value
            num_prod = num_prod + 1
        Loop
        
        'Pula se os dados j� foram enviados
        If sh_capa.Cells(lin_bd + i, "J").Value = "Enviado" Then
            GoTo proximo
        End If
        
        'Encontra o e-mail do respons�vel na tabela de e-mails
        k = 0
        Do While sh_capa.Cells(lin_bd + k, "M").Value <> ""
            'Se nome da tabela bd for igual a o nome da tabela de e-mails...
            If sh_capa.Cells(lin_bd + k, "M").Value = sh_capa.Cells(lin_bd + i, "B").Value Then
                'Configura a vari�vel destino com o e-mail desejado
                If sh_capa.Cells(lin_bd + k, "O").Value <> "x" Then
                    GoTo proximo
                End If
                destino = sh_capa.Cells(lin_bd + k, "N").Value
                Exit Do
            End If
            k = k + 1
        Loop
        
        'Verifica se a tabela est� vazia
        k = 0
        Do While sh_layout.Cells(lin_tabela + k, "B").Value <> ""
            k = k + 1
        Loop
        
        'Se a tabela n�o estiver vazia, apagar os dados antigos para que os novos
        'possam ser copiados
        sh_layout.Select
        If k > 1 Then
            Rows(lin_tabela & ":" & lin_tabela + k - 2).Delete shift:=xlUp
            Rows(lin_tabela & ":" & lin_tabela).ClearContents
        Else
            Rows(lin_tabela & ":" & lin_tabela).ClearContents
        End If
        
        'Cria as linhas com a mesma quantidade do n�mero de produtos
        For k = 1 To num_prod - 1
            Rows(lin_tabela & ":" & lin_tabela).Insert shift:=down, copyOrigin:=xlFormatFromRightOrBelow
        Next k
        
        'Copia para o layout
        For k = 0 To num_prod - 1
            For l = 0 To 6
                sh_layout.Cells(lin_tabela + k, 2 + l).Value = sh_capa.Cells(lin_bd + i + k, 2 + l).Value
            Next l
        Next k
        
        'Nome da loja
        sh_layout.Cells(4, "C").Value = sh_layout.Cells(lin_tabela, "B").Value & ","
        
        'Cria o e-mail a ser enviado
        Set outMail = outApp.CreateItem(0)
        
        'Enviar atrav�s de uma caixa gen�rica de e-mail
        'outMail.SentOnBehalfOfName = "teste@gmail.com"
        
        'Assunto o e-mail
        outMail.Subject = sh_capa.Cells(8, "K").Value
        'Destinat�rio
        outMail.To = destino
        'C�pia
        outMail.CC = sh_capa.Cells(9, "K").Value
        'C�pia oculta
        outMail.BCC = sh_capa.Cells(10, "K").Value
        
        'Anexa um arquivo
        'outMail.Attachments.Add "C:\Users\kevin\Desktop\VBA - Expertise\Sem t�tulo.jpg"
        
        'N�mero total de linhas da aba layout
        lin_layout = sh_layout.Cells(Rows.Count, "B").End(xlUp).Row
        
        'Seleciona a �rea a ser enviada por e-mail
        Set rng = sh_layout.Range("B2:I" & lin_layout).SpecialCells(xlCellTypeVisible)
        'Copia para o corpo do e-mail a �rea desejada atrav�s do m�todo rangetoHTML
        outMail.HTMLBody = rangetoHTML(rng, nome_layout)
        
        'Exibe para o usu�rio o e-mail
        outMail.Display
        'Deleta o e-mail da caixa de enviados se for true
        outMail.DeleteAfterSubmit = False
        'Envia o e-mail
        outMail.Send
        
        'Configura a vari�vel do e-mail como vazia para receber o pr�ximo e-mail
        Set outMail = Nothing
        
        'Preenche coluna de status
        For k = 0 To num_prod - 1
            sh_capa.Cells(lin_bd + i + k, "H").Value = "Enviado"
        Next k
        
        'Conta quantos e-mails foram enviados
        cont_email = cont_email + 1

proximo:
        
        'Pr�ximo nome
        i = i + num_prod
        
    Loop
    
    sh_capa.Select
    
    'Mensagem exibida para o usu�rio
    MsgBox "E-mails enviados: " & Format(cont_email, "000")
    
    'Volta a exibir alertas
    Application.DisplayAlerts = True
    'Volta a atualizar a tela
    Application.ScreenUpdating = True
    
End Sub

